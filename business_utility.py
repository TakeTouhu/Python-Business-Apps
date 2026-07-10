# -*- coding: utf-8 -*-
"""業務支援ユーティリティ

タブ1: フォルダー作成
タブ2: ファイル移動
タブ3: Outlook 未返信メール確認(スレッド単位)
タブ4: AI タスク抽出(Claude API)

必要パッケージ: pywin32, anthropic
    pip install pywin32 anthropic
AI タスク抽出には環境変数 ANTHROPIC_API_KEY の設定が必要です。
"""

import os
import shutil
import datetime
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import win32com.client
import pythoncom  # COM初期化用

# 拡張子リスト
extensions_all = ['.pdf', '.xlsx', '.xlsm', '.xls', '.pptx', '.mp4', '.png', '.txt', '.zip']

# AI タスク抽出で一度に処理するメールの上限(トークン量の暴走防止)
MAX_EMAILS_FOR_AI = 50


# ---------- 共通ヘルパー ----------
def ui(func, *args, **kwargs):
    """ワーカースレッドから GUI を安全に更新する(tkinter はスレッド非対応のため)。"""
    root.after(0, lambda: func(*args, **kwargs))


def to_naive(dt):
    """Outlook の ReceivedTime(タイムゾーン付き)を naive datetime に変換する。

    タイムゾーン付きのまま datetime.now() と比較すると TypeError になる。
    """
    try:
        return dt.replace(tzinfo=None)
    except Exception:
        return datetime.datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


def restrict_by_received(items, cutoff):
    """Items を受信日時でフィルタする。全件走査を避けるため Restrict を使う。"""
    items.Sort("[ReceivedTime]", True)
    filter_str = "[ReceivedTime] >= '" + cutoff.strftime("%m/%d/%Y %I:%M %p") + "'"
    return items.Restrict(filter_str)


def get_my_addresses(namespace):
    """自分のアドレス一覧を取得する。

    Exchange 環境では SenderEmailAddress が X.500 形式(/O=...)になるため、
    SMTP アドレスと Exchange DN の両方を集めておく。
    """
    addrs = set()
    try:
        for i in range(1, namespace.Accounts.Count + 1):
            try:
                smtp = namespace.Accounts.Item(i).SmtpAddress
                if smtp:
                    addrs.add(smtp.lower())
            except Exception:
                pass
    except Exception:
        pass
    try:
        current_user = namespace.CurrentUser
        entry = current_user.AddressEntry
        if entry.Address:
            addrs.add(entry.Address.lower())  # Exchange DN の場合あり
        if entry.Type == "EX":
            ex_user = entry.GetExchangeUser()
            if ex_user and ex_user.PrimarySmtpAddress:
                addrs.add(ex_user.PrimarySmtpAddress.lower())
    except Exception:
        pass
    return addrs


def is_me(sender, my_addrs):
    sender = (sender or "").lower()
    if not sender:
        return False
    return any(a and (a in sender or sender in a) for a in my_addrs)


# ---------- タブ1:フォルダー作成 ----------
def create_folders():
    base_path = folder_path_entry.get().strip()
    if not base_path:
        messagebox.showwarning("入力エラー", "保存先のフォルダーパスを入力してください。")
        return
    if not os.path.isdir(base_path):
        messagebox.showwarning("入力エラー", "保存先のフォルダーパスが存在しません。")
        return

    folders = ["1.得意先", "2.仕入先"]
    try:
        for name in folders:
            os.makedirs(os.path.join(base_path, name), exist_ok=True)
        messagebox.showinfo("完了", "フォルダーを作成しました。")
    except Exception as e:
        messagebox.showerror("エラー", f"フォルダー作成に失敗しました:\n{e}")


def browse_folder_1():
    selected = filedialog.askdirectory()
    if selected:
        folder_path_entry.delete(0, tk.END)
        folder_path_entry.insert(0, selected)


# ---------- タブ2:ファイル移動 ----------
def browse_source_folder():
    selected = filedialog.askdirectory()
    if selected:
        source_entry.delete(0, tk.END)
        source_entry.insert(0, selected)


def browse_target_folder():
    selected = filedialog.askdirectory()
    if selected:
        target_entry.delete(0, tk.END)
        target_entry.insert(0, selected)


def unique_destination(directory, file_name):
    """移動先に同名ファイルがある場合は _1, _2 ... を付けて衝突を回避する。"""
    base, ext = os.path.splitext(file_name)
    candidate = os.path.join(directory, file_name)
    i = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{base}_{i}{ext}")
        i += 1
    return candidate


def move_files():
    source = source_entry.get().strip()
    target = target_entry.get().strip()
    selected_exts = [ext for ext, var in ext_vars.items() if var.get()]

    if not selected_exts:
        messagebox.showwarning("入力エラー", "移動するファイル形式を1つ以上選択してください。")
        return
    if not os.path.isdir(source) or not os.path.isdir(target):
        messagebox.showwarning("エラー", "移動元または移動先フォルダーのパスが無効です。")
        return
    if os.path.abspath(source) == os.path.abspath(target):
        messagebox.showwarning("エラー", "移動元と移動先が同じフォルダーです。")
        return

    moved_count = 0
    errors = []
    for file_name in os.listdir(source):
        src_path = os.path.join(source, file_name)
        if not os.path.isfile(src_path):
            continue  # サブフォルダー等はスキップ
        ext = os.path.splitext(file_name)[1].lower()
        if ext not in selected_exts:
            continue
        try:
            shutil.move(src_path, unique_destination(target, file_name))
            moved_count += 1
        except Exception as e:
            errors.append(f"{file_name}: {e}")

    if errors:
        messagebox.showwarning(
            "一部エラー",
            f"{moved_count} 件を移動しました。\n以下のファイルは移動できませんでした:\n" + "\n".join(errors[:10]),
        )
    else:
        messagebox.showinfo("完了", f"{moved_count} 件のファイルを移動しました。")


# ---------- タブ3:Outlook未返信抽出(Conversation) ----------
def check_outlook_unreplied():
    """ボタンから呼ばれる。GUI を固まらせないよう別スレッドで実行する。"""
    unreplied_button.config(state="disabled")
    result_text.delete("1.0", tk.END)
    result_text.insert(tk.END, "🔍 Outlook を確認中...\n")
    thread = threading.Thread(target=check_outlook_unreplied_worker, daemon=True)
    thread.start()


def check_outlook_unreplied_worker():
    pythoncom.CoInitialize()  # スレッドごとに COM 初期化が必要
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # olFolderInbox

        two_weeks_ago = datetime.datetime.now() - datetime.timedelta(days=14)
        # Restrict で期間フィルタ(全メール走査を避ける・タイムゾーン比較問題も回避)
        recent = restrict_by_received(inbox.Items, two_weeks_ago)

        my_addrs = get_my_addresses(namespace)

        checked_conversations = set()
        unreplied = []

        for message in recent:
            try:
                if message.Class != 43:  # olMail 以外(会議依頼等)はスキップ
                    continue

                received = to_naive(message.ReceivedTime)

                conv_id = message.ConversationID
                if not conv_id or conv_id in checked_conversations:
                    continue
                checked_conversations.add(conv_id)

                # 自分が送ったメールが受信トレイにある場合はスキップ
                if is_me(message.SenderEmailAddress, my_addrs):
                    continue

                conversation = message.GetConversation()
                if not conversation:
                    continue

                table = conversation.GetTable()
                table.ResetColumns()
                table.Columns.Add("SenderEmailAddress")
                table.Columns.Add("ReceivedTime")

                has_reply = False
                while not table.EndOfTable:
                    row = table.GetNextRow()
                    sender = row["SenderEmailAddress"]
                    if is_me(sender, my_addrs):
                        # 受信より後の自分のメールがあれば「返信済み」
                        try:
                            if to_naive(row["ReceivedTime"]) >= received:
                                has_reply = True
                                break
                        except Exception:
                            has_reply = True
                            break

                if not has_reply:
                    unreplied.append(
                        f"[{received.strftime('%Y-%m-%d %H:%M')}] {message.SenderName}:{message.Subject}"
                    )
            except Exception as e:
                print(f"スキップされたメール: {e}")
                continue

        def show_result():
            result_text.delete("1.0", tk.END)
            if unreplied:
                result_text.insert(tk.END, f"📩 未返信スレッド {len(unreplied)} 件\n" + "=" * 60 + "\n")
                for line in unreplied:
                    result_text.insert(tk.END, line + "\n")
            else:
                result_text.insert(tk.END, "未返信のメールは見つかりませんでした。")
            unreplied_button.config(state="normal")

        ui(show_result)

    except Exception as e:
        err = str(e)

        def show_error():
            result_text.delete("1.0", tk.END)
            unreplied_button.config(state="normal")
            messagebox.showerror("エラー", f"Outlookとの連携に失敗しました:\n{err}")

        ui(show_error)
    finally:
        pythoncom.CoUninitialize()


# ---------- タブ4:AIタスク抽出 ----------
def get_recent_emails(hours=24):
    """過去 hours 時間のメールを取得(呼び出し元スレッドで COM 初期化済みであること)。"""
    namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)

    cutoff = datetime.datetime.now() - datetime.timedelta(hours=hours)
    recent = restrict_by_received(inbox.Items, cutoff)

    recent_emails = []
    for message in recent:
        try:
            if message.Class != 43:
                continue
            received = to_naive(message.ReceivedTime)
            recent_emails.append({
                'subject': message.Subject or "",
                'sender': message.SenderName or "",
                'body': (message.Body or "")[:1000],
                'received': received.strftime('%Y-%m-%d %H:%M'),
            })
            if len(recent_emails) >= MAX_EMAILS_FOR_AI:
                break
        except Exception as e:
            print(f"メール処理エラー: {e}")
            continue
    return recent_emails


# 構造化出力スキーマ:必ずこの形式の JSON が返るため、パース失敗が起きない
TASK_SCHEMA = {
    "type": "object",
    "properties": {
        "tasks": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "task": {"type": "string"},
                    "priority": {"type": "string", "enum": ["高", "中", "低"]},
                    "deadline": {"type": "string"},
                    "from": {"type": "string"},
                    "email_subject": {"type": "string"},
                },
                "required": ["task", "priority", "deadline", "from", "email_subject"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["tasks"],
    "additionalProperties": False,
}


def extract_tasks_with_ai(emails, status_callback):
    """Claude API でタスクを抽出する。"""
    try:
        import anthropic
    except ImportError:
        status_callback("⚠️ anthropicパッケージがインストールされていません。\npip install anthropic を実行してください。")
        return []

    email_text = ""
    for i, email in enumerate(emails, 1):
        email_text += f"\n--- メール {i} ---\n"
        email_text += f"件名: {email['subject']}\n"
        email_text += f"送信者: {email['sender']}\n"
        email_text += f"受信時刻: {email['received']}\n"
        email_text += f"本文:\n{email['body']}\n"

    prompt = f"""以下のメールから、行動が必要なタスクを抽出してください。

タスクの例:
- 返信が必要なメール
- 書類の提出依頼
- 会議の出席確認
- レビュー依頼
- 情報提供の依頼

各タスクの deadline は期限が明記されていれば記載し、なければ空文字にしてください。
広告・通知など行動不要なメールからはタスクを作らないでください。

メール内容:
{email_text}"""

    try:
        client = anthropic.Anthropic()
        response = client.messages.create(
            model="claude-opus-4-8",
            max_tokens=8192,
            output_config={"format": {"type": "json_schema", "schema": TASK_SCHEMA}},
            messages=[{"role": "user", "content": prompt}],
        )

        if response.stop_reason == "refusal":
            status_callback("⚠️ AIがリクエストを拒否しました。")
            return []
        if response.stop_reason == "max_tokens":
            status_callback("⚠️ 出力が上限に達しました。対象メール数を減らしてください。")
            return []

        text = next((b.text for b in response.content if b.type == "text"), "")
        return json.loads(text)["tasks"]

    except anthropic.AuthenticationError:
        status_callback("⚠️ APIキーが無効です。環境変数 ANTHROPIC_API_KEY を確認してください。")
        return []
    except anthropic.APIConnectionError:
        status_callback("⚠️ ネットワークエラーです。接続を確認して再試行してください。")
        return []
    except Exception as e:
        status_callback(f"⚠️ AI抽出エラー: {e}")
        return []


def extract_tasks_worker():
    """タスク抽出のワーカースレッド本体。GUI 更新はすべて ui() 経由で行う。"""

    def update_status(msg):
        def _append():
            task_result_text.insert(tk.END, msg + "\n")
            task_result_text.see(tk.END)
        ui(_append)

    pythoncom.CoInitialize()  # このスレッドで COM を使うため
    try:
        emails = get_recent_emails(hours=24)
        update_status(f"✅ {len(emails)}件のメールを取得しました\n")

        if not emails:
            update_status("過去24時間に新しいメールはありません")
            return

        update_status("🤖 AIでタスクを抽出中...\n")
        tasks = extract_tasks_with_ai(emails, update_status)

        if not tasks:
            update_status("⚠️ タスクが見つかりませんでした")
            return

        update_status(f"\n📋 抽出されたタスク ({len(tasks)}件):\n")
        update_status("=" * 60 + "\n")

        for i, task in enumerate(tasks, 1):
            update_status(f"\n【タスク {i}】")
            update_status(f"  内容: {task['task']}")
            update_status(f"  優先度: {task['priority']}")
            update_status(f"  期限: {task.get('deadline') or '指定なし'}")
            update_status(f"  送信者: {task['from']}")
            update_status(f"  元のメール: {task['email_subject']}")
            update_status("-" * 60)

        # タスクをJSONファイルに保存
        try:
            output_file = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                f"tasks_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            )
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(tasks, f, ensure_ascii=False, indent=2)
            update_status(f"\n💾 タスクを {output_file} に保存しました")
        except Exception as e:
            update_status(f"\n⚠️ ファイル保存に失敗しました: {e}")

    except Exception as e:
        update_status(f"❌ エラー: {e}")
    finally:
        pythoncom.CoUninitialize()
        ui(lambda: extract_button.config(state="normal"))


def extract_tasks():
    """タスク抽出を開始"""
    extract_button.config(state="disabled")
    task_result_text.delete("1.0", tk.END)
    task_result_text.insert(tk.END, "🔍 メールを取得中...\n")
    thread = threading.Thread(target=extract_tasks_worker, daemon=True)
    thread.start()


# ---------- GUI構築 ----------
root = tk.Tk()
root.title("業務支援ユーティリティ")
root.geometry("700x550")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# ----- タブ1:フォルダー作成 -----
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="フォルダー作成")

tk.Label(tab1, text="保存先フォルダーを指定:").pack(pady=10)
frame1 = tk.Frame(tab1)
frame1.pack()
folder_path_entry = tk.Entry(frame1, width=45)
folder_path_entry.pack(side=tk.LEFT, padx=5)
tk.Button(frame1, text="参照", command=browse_folder_1).pack(side=tk.LEFT)
tk.Button(tab1, text="得意先 / 仕入先 フォルダーを生成", command=create_folders, bg="lightblue", width=40).pack(pady=20)

# ----- タブ2:ファイル移動 -----
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="ファイル移動")

tk.Label(tab2, text="移動元フォルダーを指定:").pack(pady=(10, 5))
frame2 = tk.Frame(tab2)
frame2.pack()
source_entry = tk.Entry(frame2, width=45)
source_entry.pack(side=tk.LEFT, padx=5)
tk.Button(frame2, text="参照", command=browse_source_folder).pack(side=tk.LEFT)

tk.Label(tab2, text="移動先フォルダーを指定:").pack(pady=(10, 5))
frame3 = tk.Frame(tab2)
frame3.pack()
target_entry = tk.Entry(frame3, width=45)
target_entry.pack(side=tk.LEFT, padx=5)
tk.Button(frame3, text="参照", command=browse_target_folder).pack(side=tk.LEFT)

tk.Label(tab2, text="移動対象の拡張子:").pack(pady=(10, 5))
ext_vars = {}
frame_exts = tk.Frame(tab2)
frame_exts.pack()
for ext in extensions_all:
    var = tk.BooleanVar(value=True)
    cb = tk.Checkbutton(frame_exts, text=ext, variable=var)
    cb.pack(side=tk.LEFT, padx=5)
    ext_vars[ext] = var

tk.Button(tab2, text="選択した形式のファイルを移動", command=move_files, bg="lightgreen", width=40).pack(pady=20)

# ----- タブ3:Outlook未返信メール(Conversation単位) -----
tab3 = ttk.Frame(notebook)
notebook.add(tab3, text="未返信確認")

tk.Label(tab3, text="過去2週間の未返信メール(スレッド単位)を表示します").pack(pady=10)
unreplied_button = tk.Button(tab3, text="未返信メールを抽出", command=check_outlook_unreplied,
                             bg="lightyellow", width=30)
unreplied_button.pack(pady=5)

result_text = scrolledtext.ScrolledText(tab3, width=80, height=22)
result_text.pack(pady=10, padx=10)

# ----- タブ4:AIタスク抽出 -----
tab4 = ttk.Frame(notebook)
notebook.add(tab4, text="AIタスク抽出")

tk.Label(tab4, text="過去24時間のメールからAIが自動的にタスクを抽出します",
         font=("", 10)).pack(pady=10)

info_frame = tk.Frame(tab4)
info_frame.pack(pady=5)
tk.Label(info_frame, text="💡 事前準備: pip install anthropic と 環境変数 ANTHROPIC_API_KEY の設定",
         fg="blue", font=("", 9)).pack()

extract_button = tk.Button(tab4, text="🤖 AIでタスクを抽出", command=extract_tasks,
                           bg="lightcoral", width=30, font=("", 10, "bold"))
extract_button.pack(pady=10)

task_result_text = scrolledtext.ScrolledText(tab4, width=80, height=20,
                                             font=("Consolas", 9))
task_result_text.pack(pady=10, padx=10)

root.mainloop()
