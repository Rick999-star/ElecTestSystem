/**
 * デバッグ用: API接続テスト関数
 * GASエディタの上部バーから「testGeminiConnection」を選択して「実行」してください。
 */
function testGeminiConnection() {
    const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEY_GEMINI_API_KEY);

    if (!apiKey) {
        console.error("【エラー】APIキーが設定されていません。スクリプトプロパティを確認してください。");
        return;
    }

    // 利用可能なモデル一覧を取得するAPI
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;

    try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const code = response.getResponseCode();
        const text = response.getContentText();

        console.log(`レスポンスコード: ${code}`);

        if (code === 200) {
            const data = JSON.parse(text);
            console.log("【成功】接続できました。利用可能なモデル一覧:");
            data.models.forEach(m => {
                // "models/gemini-1.5-flash" のような形式で出力されます
                if (m.name.includes('gemini')) {
                    console.log(` - ${m.name}`);
                }
            });
        } else {
            console.error(`【失敗】エラーが返ってきました: ${text}`);
            console.log("ヒント: エラー400ならAPIキーが無効、403なら権限不足、404ならURL間違いの可能性があります");
        }

    } catch (e) {
        console.error("【通信エラー】: " + e.toString());
    }
}
