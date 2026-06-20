from app import create_app

app = create_app()

if __name__ == "__main__":
    # threaded=True: スライド画像生成を並行実行できるようにする
    # use_reloader=False / debug=False: リローダーの子プロセス化を防ぐ。
    #   （子プロセス化すると macOS の TCC 責任プロセス継承が切れ、iCloud Drive へ
    #     アクセスできなくなるため。単一プロセスで起動する）
    # port=5070: 5060 は SIP の予約ポートで Chrome が ERR_UNSAFE_PORT として
    #   接続を拒否するため、安全なポートを使う
    app.run(debug=False, use_reloader=False, port=5070, threaded=True)
