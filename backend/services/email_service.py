"""Outlook email draft creation via COM automation.

ローカルのOutlookアプリケーションを使用してメールの下書きを作成・表示する。
ユーザーは内容を確認後、自分で送信ボタンを押す。
"""

import platform


def create_outlook_draft(
    to_email: str,
    subject: str,
    body: str,
) -> dict:
    """Open Outlook with a pre-filled draft email."""
    if platform.system() != "Windows":
        raise RuntimeError("Outlook連携はWindowsでのみ利用可能です")

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32がインストールされていません。\n"
            "pip install pywin32 を実行してください。"
        )

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body
        mail.Display()  # 下書きを表示（ユーザーが送信ボタンを押す）
    except Exception as e:
        raise RuntimeError(
            f"Outlookの起動に失敗しました: {e}\n"
            "Outlookがインストールされているか確認してください。"
        )
    finally:
        pythoncom.CoUninitialize()

    return {"status": "ok", "message": "Outlookでメールの下書きを開きました"}
