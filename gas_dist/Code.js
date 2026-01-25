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
        const header = ['ID', 'Text', 'Image URL', 'Points'];
        sheet.appendRow(header);

        // データ書き込み
        if (questions && questions.length > 0) {
            const rows = questions.map(q => [
                q.id,
                q.text,
                q.imageUrl || '',
                q.points
            ]);
            sheet.getRange(2, 1, rows.length, 4).setValues(rows);
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

        const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

        // オブジェクト配列に変換
        return data.map(row => ({
            id: row[0],
            text: row[1],
            imageUrl: row[2],
            points: Number(row[3])
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

【学生の回答】
${studentAnswer}

【採点フォーマット（厳守）】
以下のJSON形式のみを出力してください。Markdownのコードブロックは不要です。
{"score": 数値(0-${q.points}), "reason": "採点理由とフィードバック（100文字程度）"}
`;

            const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
            const payload = {
                contents: [{ parts: [{ text: prompt }] }]
            };

            const options = {
                'method': 'post',
                'contentType': 'application/json',
                'payload': JSON.stringify(payload)
            };

            const response = UrlFetchApp.fetch(url, options);
            const data = JSON.parse(response.getContentText());
            const text = data.candidates[0].content.parts[0].text;

            // JSONパース (GeminiがたまにMarkdownブロックを含めるため除去)
            const jsonStr = text.replace(/```json/g, '').replace(/```/g, '').trim();
            const assessment = JSON.parse(jsonStr);

            results.push({
                questionId: q.id,
                score: assessment.score,
                reason: assessment.reason
            });

        } catch (apiError) {
            console.error("Gemini API Error for Q " + q.id, apiError);
            results.push({ questionId: q.id, score: 0, reason: "採点エラー: AI通信に失敗しました" });
        }
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
