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
        const header = ['ID', 'Text', 'Image URL', 'Points', 'Criteria', 'SubQuestionsJSON', 'ModelAnswer'];
        sheet.appendRow(header);

        // データ書き込み
        if (questions && questions.length > 0) {
            const rows = questions.map(q => [
                q.id,
                q.text,
                q.imageUrl || '',
                q.points,
                q.criteria || '',
                q.subQuestions ? JSON.stringify(q.subQuestions) : '', // 小問をJSONとして保存
                q.modelAnswer || '' // Col 7: 表示用模範解答
            ]);
            sheet.getRange(2, 1, rows.length, 7).setValues(rows);
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

        // 7列目まで取得 (Col G: ModelAnswer)
        const maxCols = sheet.getMaxColumns();
        const numColsToGet = Math.min(7, maxCols); // 最大でも7列

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

            // Col 7 (index 6) が modelAnswer
            const modelAnswer = (row.length >= 7) ? row[6] : '';

            return {
                id: row[0],
                text: row[1],
                imageUrl: row[2],
                points: Number(row[3]),
                criteria: row[4] || '',
                subQuestions: subQuestions,
                modelAnswer: modelAnswer
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
        const errorMsg = e.toString();
        // ユーザーにわかりやすいエラーメッセージを返す
        return { success: false, message: '送信エラー: ' + errorMsg };
    }
}

/**
 * Gemini APIと通信して採点を行う内部関数
 */
/**
 * Gemini APIと通信して採点を行う内部関数 (一括採点版)
 */
function _gradeWithGemini(questions, answers) {
    const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEY_GEMINI_API_KEY);

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

    // プロンプト構築用に問題をフラット化してリスト作成
    const problemList = [];
    questions.forEach(q => {
        if (q.subQuestions && q.subQuestions.length > 0) {
            q.subQuestions.forEach(sq => {
                problemList.push({
                    type: 'sub',
                    qId: q.id,
                    sqId: sq.id,
                    text: `[親問題]: ${q.text}\n[小問題]: ${sq.text}`,
                    points: sq.points,
                    criteria: sq.criteria || '特になし',
                    studentAnswer: (answers[q.id] && answers[q.id][sq.id]) || ""
                });
            });
        } else {
            problemList.push({
                type: 'normal',
                qId: q.id,
                sqId: null,
                text: q.text,
                points: q.points,
                criteria: q.criteria || '特になし',
                studentAnswer: answers[q.id] || ""
            });
        }
    });

    if (problemList.length === 0) return [];

    // 一括送信だと429エラー(Resource exhausted)になるため、分割処理(チャンク化)を行う
    const CHUNK_SIZE = 5; // 5問ずつ処理
    let allResults = [];

    // チャンクごとの処理ループ
    for (let i = 0; i < problemList.length; i += CHUNK_SIZE) {
        const chunk = problemList.slice(i, i + CHUNK_SIZE);

        let promptText = `
あなたは電気工学の専門家かつ厳格な採点者です。以下の試験問題に対する学生の回答を一括で採点してください。

【採点対象リスト】
`;

        chunk.forEach((p, index) => {
            const ans = typeof p.studentAnswer === 'string' ? p.studentAnswer : JSON.stringify(p.studentAnswer);
            promptText += `
---
No.${index + 1}
[問題ID: ${p.qId}${p.sqId ? '_' + p.sqId : ''}]
問題文: ${p.text}
配点: ${p.points}点
採点基準: ${p.criteria}
学生の回答: ${ans || '(未回答)'}
`;
        });

        promptText += `
---

【採点フォーマット（厳守）】
以下のJSON配列形式のみを出力してください。Markdownのコードブロックは不要です。
必ず "No.1" から "No.${chunk.length}" までの全ての採点結果を含めてください。

[
  { "index": 0, "score": 数値, "reason": "短いフィードバック" },
  { "index": 1, "score": 数値, "reason": "短いフィードバック" },
  ...
]

※ indexは0始まり(No.1に対応)で、入力リストの順序と一致させてください。
※ 未回答の場合は0点としてください。
※ 理由(reason)は学生への直接的なフィードバックとして適切かつ簡潔な日本語で記述してください。
`;

        try {
            // レート制限回避のため、2回目以降は少し待機
            if (i > 0) Utilities.sleep(1000);

            const aiResponse = _callGeminiApi(apiKey, promptText);
            let chunkResults = [];

            if (!Array.isArray(aiResponse)) {
                // 単一オブジェクト救済
                chunkResults = Array.isArray(aiResponse) ? aiResponse : [aiResponse];
            } else {
                chunkResults = aiResponse;
            }

            // チャンク内インデックスを元にマッピング
            const mappedChunk = chunkResults.map(r => {
                const idx = r.index;
                // インデックス範囲チェック
                if (typeof idx !== 'number' || idx < 0 || idx >= chunk.length) return null;

                const originalParam = chunk[idx];
                return {
                    questionId: originalParam.qId,
                    subQuestionId: originalParam.sqId,
                    score: Number(r.score) || 0,
                    reason: r.reason || ''
                };
            }).filter(Boolean);

            allResults = allResults.concat(mappedChunk);

        } catch (apiError) {
            console.error(`Gemini API Batch Error (Chunk ${i / CHUNK_SIZE + 1})`, apiError);
            // エラー時はこのチャンク分を0点で埋める
            const fallback = chunk.map(p => ({
                questionId: p.qId,
                subQuestionId: p.sqId,
                score: 0,
                reason: "採点システムエラー: " + apiError.message
            }));
            allResults = allResults.concat(fallback);
        }
    }

    return allResults;
}

/**
 * Gemini API呼び出し共通化 (リトライ処理付き)
 */
function _callGeminiApi(apiKey, prompt) {
    // ユーザー環境で利用可能な最新モデル (Gemini 3.0 Flash Preview) に切り替え
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;
    const payload = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
            response_mime_type: "application/json"
        }
    };

    const options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };

    const MAX_RETRIES = 3;
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
        return {
            success: false,
            message: "【エラー】APIキーが設定されていません。GASの「プロジェクトの設定」>「スクリプトプロパティ」を確認してください。"
        };
    }

    // 利用可能なモデル一覧を取得するAPI
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;

    try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const code = response.getResponseCode();
        const text = response.getContentText();

        if (code === 200) {
            const data = JSON.parse(text);
            const models = data.models
                .filter(m => m.name.includes('gemini'))
                .map(m => m.name)
                .join(', ');
            return {
                success: true,
                message: "接続成功！利用可能モデル: " + models
            };
        } else {
            return {
                success: false,
                message: `エラー (${code}): ${text}`
            };
        }

    } catch (e) {
        return {
            success: false,
            message: "通信エラー: " + e.toString()
        };
    }
}

/**
 * デバッグ用: 採点ロジック単体テスト
 */
function testGrading() {
    const dummyQuestions = [{
        id: 'debug_q1',
        text: '電流の単位は何か？記号で答えなさい。',
        points: 10,
        criteria: '正解は「A」または「アンペア」'
    }];
    const dummyAnswers = { 'debug_q1': 'A' };

    try {
        const start = new Date();
        // testGradingの中で submitAnswers を呼んでみる（サーバー内部での呼び出しテスト）
        const results = submitAnswers(dummyAnswers);
        const end = new Date();
        const duration = (end - start) / 1000;

        return {
            success: true,
            message: `採点テスト(Internal submitAnswers) 成功 (${duration}秒)`,
            details: results
        };
    } catch (e) {
        return {
            success: false,
            message: "採点テスト失敗: " + e.toString()
        };
    }
}
