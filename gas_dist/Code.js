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
        const header = ['ID', 'Text', 'Image URL', 'Points', 'Criteria', 'SubQuestionsJSON'];
        sheet.appendRow(header);

        // データ書き込み
        if (questions && questions.length > 0) {
            const rows = questions.map(q => [
                q.id,
                q.text,
                q.imageUrl || '',
                q.points,
                q.criteria || '',
                q.subQuestions ? JSON.stringify(q.subQuestions) : '' // 小問をJSONとして保存
            ]);
            sheet.getRange(2, 1, rows.length, 6).setValues(rows);
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

        // 6列目まで取得するように変更 (Col F: SubQuestionsJSON)
        // 既存シートが5列しかない場合のエラー回避のため、getRangeの列数はシートの最大列数などを考慮するのが安全だが、
        // ここでは新規保存で列が増える前提とする。読み込み時に列が足りない場合は空文字が返ることを期待。
        // SpreadsheetAppでは指定範囲がシート範囲外だとエラーになる可能性があるが、データがある範囲(getDataRange)を使う手もある。
        // ここでは安全のため getDataRange を使いつつ、必要な部分をマップするアプローチに変更するか、
        // あるいは6列固定で取得する。列が存在しない場合はエラーになるので、列数チェックを入れる。

        const maxCols = sheet.getMaxColumns();
        const numColsToGet = Math.min(6, maxCols); // 最大でも6列

        const data = sheet.getRange(2, 1, lastRow - 1, numColsToGet).getValues();

        // オブジェクト配列に変換
        return data.map(row => {
            let subQuestions = [];
            if (row.length >= 6 && row[5]) {
                try {
                    subQuestions = JSON.parse(row[5]);
                } catch (e) {
                    console.warn("Failed to parse subQuestions JSON", e);
                }
            }

            return {
                id: row[0],
                text: row[1],
                imageUrl: row[2],
                points: Number(row[3]),
                criteria: row[4] || '',
                subQuestions: subQuestions
            };
        });
    } catch (e) {
        console.error(e);
        // エラー時は空リストを返す（またはエラーをスロー）
        return [];
    }
}

/**
 * 回答の送信と採点
 * @param {Object} answers - { questionId: answerText, ... } または { questionId: { subId: answerText, ... } }
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
        return questions.flatMap(q => {
            if (q.subQuestions && q.subQuestions.length > 0) {
                return q.subQuestions.map(sq => ({
                    questionId: q.id,
                    subQuestionId: sq.id,
                    score: Math.floor(sq.points * 0.8),
                    reason: "(Mock) Sub-question graded."
                }));
            } else {
                return [{
                    questionId: q.id,
                    score: Math.floor(q.points * 0.8),
                    reason: "(Mock) Graded."
                }];
            }
        });
    }

    for (const q of questions) {
        // 小問がある場合の処理
        if (q.subQuestions && q.subQuestions.length > 0) {
            const result = _gradeSubQuestions(q, answers[q.id] || {}, apiKey);
            if (result) results.push(...result);
            // APIレート制限考慮
            Utilities.sleep(1000);
            continue;
        }

        // 従来の問題（小問なし）の処理
        const studentAnswer = answers[q.id] || "";

        // 文字列でない場合(回答形式不正)への対応
        const answerText = typeof studentAnswer === 'string' ? studentAnswer : JSON.stringify(studentAnswer);

        // 空欄の場合は0点
        if (!answerText.trim()) {
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
${answerText}

【採点フォーマット（厳守）】
以下のJSON形式のみを出力してください。Markdownのコードブロックは不要です。
**注意: 数式などでバックスラッシュを使用する場合は、必ず "\\\\" (二重) にしてエスケープしてください。**
{"score": 数値(0-${q.points}), "reason": "採点理由とフィードバック（100文字程度）"}
`;

            const aiResponse = _callGeminiApi(apiKey, prompt);
            results.push({
                questionId: q.id,
                score: aiResponse.score,
                reason: aiResponse.reason
            });

        } catch (apiError) {
            console.error("Gemini API Error for Q " + q.id, apiError);
            results.push({ questionId: q.id, score: 0, reason: "採点エラー: " + apiError.toString() });
        }

        Utilities.sleep(1000);
    }

    return results;
}

/**
 * 小問付き問題の採点
 */
function _gradeSubQuestions(question, answerObj, apiKey) {
    // answerObj は k:v 形式を想定 (subId: answerText)

    // 全空欄チェック等は省略し、AIに一括で投げます

    // プロンプト構築
    let subQsText = "";
    let totalPoints = 0;

    question.subQuestions.forEach((sq, idx) => {
        subQsText += `
小問(${idx + 1}): ${sq.text} (配点: ${sq.points}点)
基準: ${sq.criteria || 'なし'}
学生の回答: ${answerObj[sq.id] || '(未回答)'}
`;
        totalPoints += Number(sq.points);
    });

    const prompt = `
あなたは電気工学の専門家かつ採点者です。以下の試験問題（複数の小問あり）に対する学生の回答を採点してください。
親問題の共通テキストがある場合は考慮してください。

【親問題テキスト】
${question.text}

【小問リストと回答】
${subQsText}

【採点フォーマット（厳守）】
以下のJSON形式のみを出力してください。Markdownのコードブロックは不要です。
配列で返してください。
[
  { "subQIndex": 0, "score": 数値, "reason": "フィードバック" },
  { "subQIndex": 1, "score": 数値, "reason": "フィードバック" },
  ...
]
※ subQIndexは0始まりのインデックスで対応させてください。
`;

    try {
        const aiResponse = _callGeminiApi(apiKey, prompt);

        // 配列であることを確認
        if (!Array.isArray(aiResponse)) {
            throw new Error("API response is not an array");
        }

        return aiResponse.map(r => {
            const sq = question.subQuestions[r.subQIndex];
            if (!sq) return null;
            return {
                questionId: question.id,
                subQuestionId: sq.id,
                score: r.score,
                reason: r.reason
            };
        }).filter(Boolean);

    } catch (e) {
        console.error("Gemini API Error for Q " + question.id, e);
        // エラー時は全小問0点
        return question.subQuestions.map(sq => ({
            questionId: question.id,
            subQuestionId: sq.id,
            score: 0,
            reason: "採点エラー: " + e.toString()
        }));
    }
}

/**
 * Gemini API呼び出し共通化 (リトライ処理付き)
 */
function _callGeminiApi(apiKey, prompt) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-001:generateContent?key=${apiKey}`;
    const payload = {
        contents: [{ parts: [{ text: prompt }] }]
    };

    const options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };

    const MAX_RETRIES = 5;
    let retryCount = 0;

    // リトライループ
    while (true) {
        let response;
        try {
            response = UrlFetchApp.fetch(url, options);
        } catch (e) {
            // ネットワークエラー等の場合
            if (retryCount >= MAX_RETRIES) throw e;
            console.warn(`通信エラー: ${e.toString()}。リトライします... (${retryCount + 1}/${MAX_RETRIES})`);
            retryCount++;
            Utilities.sleep(Math.pow(2, retryCount) * 1000);
            continue;
        }

        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        // 成功 (200 OK)
        if (responseCode === 200) {
            const data = JSON.parse(responseText);
            // 候補がない場合のガード
            if (!data.candidates || data.candidates.length === 0) {
                throw new Error(`No candidates returned. Response: ${responseText}`);
            }
            const text = data.candidates[0].content.parts[0].text;

            let jsonStr = text.replace(/```json/g, '').replace(/```/g, '').trim();

            // パース処置
            try {
                return JSON.parse(jsonStr);
            } catch (e) {
                const match = jsonStr.match(/\[[\s\S]*\]/) || jsonStr.match(/\{[\s\S]*\}/);
                if (match) {
                    try {
                        return JSON.parse(match[0].replace(/\\/g, '\\\\'));
                    } catch (e2) {
                        // Ignore
                    }
                }
                throw e;
            }
        }

        // リトライ対象: 429 (Too Many Requests) または 5xx (サーバーエラー)
        if (responseCode === 429 || (responseCode >= 500 && responseCode < 600)) {
            if (retryCount >= MAX_RETRIES) {
                throw new Error(`API Error (${responseCode}) リトライ上限到達: ${responseText}`);
            }
            console.warn(`API Error (${responseCode}): ${responseText}。リトライします... (${retryCount + 1}/${MAX_RETRIES})`);
            retryCount++;
            // 指数バックオフ: 2秒, 4秒, 8秒, 16秒, 32秒...
            Utilities.sleep(Math.pow(2, retryCount) * 1000);
            continue;
        }

        // その他のエラー (400 Bad Request, 403 Forbidden など) はリトライしない
        throw new Error(`API Error (${responseCode}): ${responseText}`);
    }
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
