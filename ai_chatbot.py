import argparse
import os
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

ROOT = Path(__file__).resolve().parent / "chatbot_site"


class ChatbotHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {format % args}")


def main():
    parser = argparse.ArgumentParser(description="AI Chatbot 静的ファイルサーバー")
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.environ.get("PORT", 8000)),
        help="使用するポート番号 (デフォルト: 8000)",
    )
    args = parser.parse_args()
    port = args.port

    if not ROOT.exists():
        print(f"エラー: 配信フォルダが見つかりません: {ROOT}")
        return

    try:
        server = ThreadingHTTPServer(("0.0.0.0", port), ChatbotHandler)
    except OSError:
        print(f"エラー: ポート {port} はすでに使われています。")
        return

    print(f"Chatbot server is running at http://127.0.0.1:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nサーバーを停止します。")
        server.shutdown()


if __name__ == "__main__":
    main()
