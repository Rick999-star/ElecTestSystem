/**
 * Code.js
 * 電気工学試験システム バックエンドロジック
 */

// 定数定義
const SCRIPT_PROP_KEY_SHEET_ID = 'SHEET_ID';
const SCRIPT_PROP_KEY_GEMINI_API_KEY = 'GEMINI_API_KEY';
const SHEET_NAME_QUESTIONS = 'Questions';
const SHEET_NAME_RESPONSES = 'Responses';

/**
 * Webアプリへのアクセス時にHTMLを返す
 */
function doGet(e) {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('ElecTest System')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * プロパティストアから設定を取得または初期化
 */
function _getSpreadsheetId() {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty(SCRIPT_PROP_KEY_SHEET_ID);
    if (!id) throw new Error('スプレッドシートIDが設定されていません。スクリプトプロパティで SHEET_ID を設定してください。');
    return id;
}

/**
 * 問題データの保存 (管理画面用)
 * @param {Array} questions - フロントエンドから送信された問題リスト
 */
function saveQuestions(questions) {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        let sheet = ss.getSheetByName(SHEET_NAME_QUESTIONS);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME_QUESTIONS);
        }

        // 既存データをクリアしてヘッダーを設定
        sheet.clear();
        const header = ['ID', 'Text', 'Image URL', 'Points', 'Criteria']; // Criteria追加
        sheet.appendRow(header);

        // データ書き込み
        if (questions && questions.length > 0) {
            const rows = questions.map(q => [
                q.id,
                q.text,
                q.imageUrl || '',
                q.points,
                q.criteria || '' // 基準がない場合は空文字
            ]);
            sheet.getRange(2, 1, rows.length, 5).setValues(rows); // 4列->5列
        }

        return { success: true, message: '問題を保存しました。' };
    } catch (e) {
        console.error(e);
        return { success: false, message: '保存エラー: ' + e.toString() };
    }
}

/**
 * 問題データの取得 (受験画面用)
 */
function getQuestions() {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        const sheet = ss.getSheetByName(SHEET_NAME_QUESTIONS);

        if (!sheet) return []; // シートがない場合は空リスト

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return []; // データがない場合

        const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // 4列->5列

        // オブジェクト配列に変換
        return data.map(row => ({
            id: row[0],
            text: row[1],
            imageUrl: row[2],
            points: Number(row[3]),
            criteria: row[4] || '' // 取得
        }));
    } catch (e) {
        console.error(e);
        // エラー時は空リストを返す（またはエラーをスロー）
        return [];
    }
}

/**
 * 回答の送信と採点
 * @param {Object} answers - { questionId: answerText, ... }
 */
function submitAnswers(answers) {
    try {
        const questions = getQuestions();

        // 1. Gemini APIによる採点 (各問題ごと)
        const gradingResults = _gradeWithGemini(questions, answers);
        const totalScore = gradingResults.reduce((sum, r) => sum + r.score, 0);

        // 2. 結果をスプレッドシートに保存
        _saveResponseLog(questions, answers, gradingResults, totalScore);

        return {
            success: true,
            totalScore: totalScore,
            results: gradingResults
        };

    } catch (e) {
        console.error(e);
        return { success: false, message: '送信エラー: ' + e.toString() };
    }
}

/**
 * Gemini APIと通信して採点を行う内部関数
 */
function _gradeWithGemini(questions, answers) {
    const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEY_GEMINI_API_KEY);

    // マッピング結果用
    const results = [];

    // APIキーがない場合はモック採点を行う（動作確認用）
    if (!apiKey) {
        return questions.map(q => ({
            questionId: q.id,
            score: Math.floor(q.points * 0.8), // 仮の点数
            reason: "（注意: Gemini APIキーが設定されていないため、モック採点結果を表示しています）"
        }));
    }

    // NOTE: クォータ制限等を考慮し、本来はForループで直列処理するか、バッチ処理APIを使うのが望ましい
    // ここではシンプルに直列で処理します。

    for (const q of questions) {
        const studentAnswer = answers[q.id] || "";

        // 空欄の場合は0点
        if (!studentAnswer.trim()) {
            results.push({ questionId: q.id, score: 0, reason: "未回答" });
            continue;
        }

        try {
            const prompt = `
あなたは電気工学の専門家かつ採点者です。以下の試験問題に対する学生の回答を採点してください。

【問題】
${q.text}
(配点: ${q.points}点)

【採点基準・模範解答】
${q.criteria ? q.criteria : '特になし（一般的な専門知識に基づいて採点してください）'}

【学生の回答】
${studentAnswer}

【採点フォーマット（厳守）】
以下のJSON形式のみを出力してください。Markdownのコードブロックは不要です。
**注意: 数式などでバックスラッシュを使用する場合は、必ず "\\\\" (二重) にしてエスケープしてください。**
{"score": 数値(0-${q.points}), "reason": "採点理由とフィードバック（100文字程度）"}
`;

            // お客様のアカウントで利用可能なモデルリストに「1.5」が存在しないため、
            // リストに存在し、かつ安定版である「gemini-2.0-flash-001」を指定します
            const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-001:generateContent?key=${apiKey}`;
            // 他の選択肢: gemini-2.5-pro, gemini-2.5-flash なども利用可能です
            const payload = {
                contents: [{ parts: [{ text: prompt }] }]
            };

            const options = {
                'method': 'post',
                'contentType': 'application/json',
                'payload': JSON.stringify(payload),
                'muteHttpExceptions': true // エラーレスポンスをハンドリングできるようにする
            };

            const response = UrlFetchApp.fetch(url, options);
            const responseCode = response.getResponseCode();
            const responseText = response.getContentText();

            if (responseCode !== 200) {
                throw new Error(`API Error (${responseCode}): ${responseText}`);
            }

            const data = JSON.parse(responseText);
            const text = data.candidates[0].content.parts[0].text;

            // JSONパース (GeminiがたまにMarkdownブロックを含めるため除去)
            // 1. Markdownのコードブロック記法 ```json ... ``` または ``` ... ``` を削除
            let jsonStr = text.replace(/```json/g, '').replace(/```/g, '').trim();

            // 2. JSONパース試行
            let assessment;
            try {
                assessment = JSON.parse(jsonStr);
            } catch (e) {
                console.warn("JSON Parse Retry: " + jsonStr);
                // パース失敗時、余計な文字が含まれている可能性があるので { } の範囲だけ抽出して再トライ
                const match = jsonStr.match(/\{[\s\S]*\}/);
                if (match) {
                    try {
                        assessment = JSON.parse(match[0]);
                    } catch (e2) {
                        // それでもダメならバックスラッシュを置換してトライ（危険だが救済措置）
                        assessment = JSON.parse(match[0].replace(/\\/g, '\\\\'));
                    }
                }
                if (!assessment) throw e; // それでもダメならエラーを投げる
            }

            results.push({
                questionId: q.id,
                score: assessment.score,
                reason: assessment.reason
            });

        } catch (apiError) {
            console.error("Gemini API Error for Q " + q.id, apiError);
            // デバッグ用に詳細なエラーを表示
            results.push({ questionId: q.id, score: 0, reason: "採点エラー: " + apiError.toString() });
        }

        // 有料プランになったため、待機時間を短縮 (5秒 -> 1秒)
        // ※完全に0にすると突発的な大量アクセスでエラーになることがあるため、安全策で少しだけ待ちます
        Utilities.sleep(1000);
    }

    return results;
}

/**
 * 回答ログをシートに保存
 */
function _saveResponseLog(questions, answers, gradingResults, totalScore) {
    const ssId = _getSpreadsheetId();
    const ss = SpreadsheetApp.openById(ssId);
    let sheet = ss.getSheetByName(SHEET_NAME_RESPONSES);
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME_RESPONSES);
        sheet.appendRow(['Timestamp', 'Total Score', 'Details (JSON)']); // Header
    }

    const detailObj = {
        answers: answers,
        grading: gradingResults
    };

    sheet.appendRow([
        new Date(),
        totalScore,
        JSON.stringify(detailObj)
    ]);
}

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
