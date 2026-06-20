from app import create_app

app = create_app()

if __name__ == "__main__":
    # threaded=True: スライド画像生成を並行実行できるようにする
    # use_reloader=False / debug=False: リローダーの子プロセス化を防ぐ。
    #   （子プロセス化すると macOS の TCC 責任プロセス継承が切れ、iCloud Drive へ
    #     アクセスできなくなるため。単一プロセスで起動する）
    app.run(debug=False, use_reloader=False, port=5060, threaded=True)
