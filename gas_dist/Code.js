/**
 * Code.js
 * 電気工学試験システム バックエンドロジック
 */

// 定数定義
const SCRIPT_PROP_KEY_SHEET_ID = 'SHEET_ID';
const SCRIPT_PROP_KEY_GEMINI_API_KEY = 'GEMINI_API_KEY';
const SHEET_NAME_QUESTIONS = 'Questions';
const SHEET_NAME_RESPONSES = 'Responses';
const SHEET_NAME_PATTERNS = 'Patterns';
const SHEET_NAME_SCORE_TABLE = '点数表';

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
 * 問題データの保存 (管理画面用) -- 公開(Deploy)処理
 * @param {Array} questions - フロントエンドから送信された問題リスト
 * @param {string} patternTitle - (Optional) 適用するパターン名
 */
function saveQuestions(questions, patternTitle) {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        let sheet = ss.getSheetByName(SHEET_NAME_QUESTIONS);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME_QUESTIONS);
        }

        // 保存されたパターン名をプロパティに記録 (Examinee表示用)
        const props = PropertiesService.getScriptProperties();
        if (patternTitle) {
            props.setProperty('CURRENT_PATTERN_TITLE', patternTitle);
        } else if (patternTitle === '') {
            props.deleteProperty('CURRENT_PATTERN_TITLE');
        }
        // patternTitleが未指定(undefined)の場合は、既存の値を維持するか、"Custom"とするか。
        // ここでは更新しない(=維持)戦略をとるが、明示的にnull/emptyが渡されたら消す。

        // 既存データをクリアしてヘッダーを設定
        sheet.clear();
        const header = ['ID', 'Text', 'Image URL', 'Points', 'Criteria', 'SubQuestionsJSON', 'ModelAnswer', 'ReferenceDiagramId'];
        sheet.appendRow(header);

        // データ書き込み
        if (questions && questions.length > 0) {
            const rows = questions.map(q => [
                q.id,
                q.text,
                q.imageUrl || '',
                q.points,
                q.criteria || '',
                q.subQuestions ? JSON.stringify(q.subQuestions) : '',
                q.modelAnswer || '',
                q.referenceDiagramId || ''
            ]);
            sheet.getRange(2, 1, rows.length, 8).setValues(rows);
        }

        SpreadsheetApp.flush();
        return { success: true, message: '問題を保存・公開しました。' };
    } catch (e) {
        console.error(e);
        return { success: false, message: '保存エラー: ' + e.toString() };
    }
}

/**
 * 現在公開中のパターン名を取得
 */
function getDeployedPatternTitle() {
    try {
        const props = PropertiesService.getScriptProperties();
        return props.getProperty('CURRENT_PATTERN_TITLE') || '';
    } catch (e) {
        console.error(e);
        return '';
    }
}

/**
 * パターンの保存
 * @param {string} title - パターン名
 * @param {Array} questions - 問題リスト
 */
function savePattern(title, questions) {
    try {
        if (!title) throw new Error('タイトルが空です。');

        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        let sheet = ss.getSheetByName(SHEET_NAME_PATTERNS);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME_PATTERNS);
            sheet.appendRow(['Title', 'QuestionsJSON', 'UpdatedAt']); // Header
        }

        const data = sheet.getDataRange().getValues();
        let rowIndex = -1;

        // 既存タイトルの検索 (2行目以降)
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]) === String(title)) {
                rowIndex = i + 1; // 1-based index
                break;
            }
        }

        const jsonStr = JSON.stringify(questions);
        const timestamp = new Date();

        if (rowIndex > 0) {
            // 上書き
            sheet.getRange(rowIndex, 2).setValue(jsonStr);
            sheet.getRange(rowIndex, 3).setValue(timestamp);
            SpreadsheetApp.flush();
            return { success: true, message: `パターン「${title}」を更新しました。` };
        } else {
            // 新規追加
            sheet.appendRow([title, jsonStr, timestamp]);
            SpreadsheetApp.flush();
            return { success: true, message: `パターン「${title}」を保存しました。` };
        }

    } catch (e) {
        console.error(e);
        return { success: false, message: '保存エラー: ' + e.toString() };
    }
}

/**
 * 保存済みパターンのリスト取得
 */
function getPatternList() {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        const sheet = ss.getSheetByName(SHEET_NAME_PATTERNS);
        if (!sheet) return [];

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        // Title列(A列)とUpdatedAt列(C列)のみ取得
        const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
        return data.map(row => ({
            title: String(row[0]),
            updatedAt: row[2] ? new Date(row[2]).toISOString() : ''
        }));
    } catch (e) {
        console.error('getPatternList Error:', e);
        return [];
    }
}

/**
 * 特定パターンの読み込み
 */
function getPattern(title) {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        const sheet = ss.getSheetByName(SHEET_NAME_PATTERNS);
        if (!sheet) throw new Error('Pattern sheet not found');

        const data = sheet.getDataRange().getValues();
        // 2行目以降を検索
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]) === String(title)) {
                const jsonStr = data[i][1];
                const questions = JSON.parse(jsonStr);
                return { success: true, questions: questions };
            }
        }
        return { success: false, message: 'パターンが見つかりませんでした。' };
    } catch (e) {
        console.error(e);
        return { success: false, message: '読み込みエラー: ' + e.toString() };
    }
}

/**
 * パターンの削除
 */
function deletePattern(title) {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        const sheet = ss.getSheetByName(SHEET_NAME_PATTERNS);
        if (!sheet) throw new Error('Pattern sheet not found');

        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]) === String(title)) {
                sheet.deleteRow(i + 1);
                return { success: true };
            }
        }
        return { success: false, message: 'パターンが見つかりませんでした。' };
    } catch (e) {
        return { success: false, message: '削除エラー: ' + e.toString() };
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

        // 8列目まで取得 (Col H: ReferenceDiagramId)
        const maxCols = sheet.getMaxColumns();
        const numColsToGet = Math.min(8, maxCols); // 最大でも8列

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
            // Col 8 (index 7) が referenceDiagramId
            const referenceDiagramId = (row.length >= 8) ? row[7] : '';

            return {
                id: row[0],
                text: row[1],
                imageUrl: row[2],
                points: Number(row[3]),
                criteria: row[4] || '',
                subQuestions: subQuestions,
                modelAnswer: modelAnswer,
                referenceDiagramId: referenceDiagramId
            };
        });
    } catch (e) {
        console.error(e);
        // エラー時は空リストを返す（またはエラーをスロー）
        return [];
    }
}

/**
 * 受験者の登録 (試験開始時)
 * @param {string} name - 受験者名
 * @param {string} patternTitle - 試験パターン名
 * @return {string} sessionId - セッションID (点数表のID)
 */
function registerCandidate(name, patternTitle) {
    try {
        if (!name) throw new Error("名前が入力されていません。");

        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        let sheet = ss.getSheetByName(SHEET_NAME_SCORE_TABLE);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME_SCORE_TABLE);
            // Header: ID | Date | Name | Pattern | Score
            sheet.appendRow(['ID', 'Timestamp', 'Name', 'Pattern', 'Score']);
        }

        const sessionId = Utilities.getUuid();
        const timestamp = new Date();

        sheet.appendRow([sessionId, timestamp, name, patternTitle || '', '']); // Score is empty initially

        SpreadsheetApp.flush();
        return sessionId;

    } catch (e) {
        console.error("registerCandidate Error:", e);
        throw e;
    }
}

/**
 * 採点前に回答を一時保存する (タイムアウト対策)
 * @param {Object} answers - 回答オブジェクト
 * @param {string} sessionId - セッションID
 */
function saveTemporaryAnswers(answers, sessionId) {
    try {
        if (!sessionId) throw new Error("Session ID is required for temporary save.");

        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        let sheet = ss.getSheetByName(SHEET_NAME_RESPONSES);

        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME_RESPONSES);
            sheet.appendRow(['Timestamp', 'Total Score', 'SessionID', 'Details (JSON)']); // Header
        }

        const detailObj = {
            status: "PENDING_GRADING", // 採点待ちフラグ
            answers: answers
        };

        // Scoreは "PENDING" として保存
        sheet.appendRow([
            new Date(),
            "PENDING",
            sessionId,
            JSON.stringify(detailObj)
        ]);

        return { success: true };
    } catch (e) {
        console.error("saveTemporaryAnswers Error:", e);
        return { success: false, message: e.toString() };
    }
}

/**
 * 回答の送信と採点
 * @param {Object} answers - 回答オブジェクト
 * @param {string} sessionId - セッションID (registerCandidateで取得)
 */
function submitAnswers(answers, sessionId) {
    try {
        const questions = getQuestions();

        // 1. Gemini APIによる採点 (各問題ごと)
        const gradingResults = _gradeWithGemini(questions, answers);
        const totalScore = gradingResults.reduce((sum, r) => sum + r.score, 0);

        // 2. 点数表の更新 (sessionIdがある場合)
        if (sessionId) {
            _updateScoreTable(sessionId, totalScore);
        }

        // 3. 詳細ログをスプレッドシートに保存 (バックアップ/詳細分析用)
        _saveResponseLog(questions, answers, gradingResults, totalScore, sessionId);

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
 * 点数表のスコア更新
 */
function _updateScoreTable(sessionId, score) {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        const sheet = ss.getSheetByName(SHEET_NAME_SCORE_TABLE);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        // 1列目(ID)を検索 (Header is row 1, data starts row 2)
        // データを走査してIDが一致する行を探す
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]) === String(sessionId)) {
                // Score is Col 5 (index 4)
                sheet.getRange(i + 1, 5).setValue(score);
                SpreadsheetApp.flush();
                return;
            }
        }
        console.warn("Session ID not found in Score Table:", sessionId);
    } catch (e) {
        console.error("Update Score Table Error:", e);
    }
}

/**
 * Gemini APIと通信して採点を行う内部関数
 */
/**
 * Gemini APIと通信して採点を行う内部関数 (一括採点版)
 */
/**
 * Gemini APIと通信して採点を行う内部関数 (並列処理版)
 */
function _gradeWithGemini(questions, answers) {
    const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEY_GEMINI_API_KEY);

    // APIキーがない場合はモック採点
    if (!apiKey) {
        return questions.flatMap(q => {
            if (q.subQuestions && q.subQuestions.length > 0) {
                return q.subQuestions.map(sq => ({
                    questionId: q.id,
                    subQuestionId: sq.id,
                    score: Math.floor(sq.points * 0.8),
                    reason: "(Mock) Sub-question graded.",
                    questionText: sq.text,
                    modelAnswer: sq.modelAnswer || ''
                }));
            } else {
                return [{
                    questionId: q.id,
                    score: Math.floor(q.points * 0.8),
                    reason: "(Mock) Graded.",
                    questionText: q.text,
                    modelAnswer: q.modelAnswer || ''
                }];
            }
        });
    }

    // 問題のフラット化
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
                    studentAnswer: (answers[q.id] && answers[q.id][sq.id]) || "",
                    modelAnswer: sq.modelAnswer || "" // Add model answer
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
                studentAnswer: answers[q.id] || "",
                modelAnswer: q.modelAnswer || "" // Add model answer
            });
        }
    });

    if (problemList.length === 0) return [];

    // チャンクサイズを10に拡大 (リクエスト数削減)
    const CHUNK_SIZE = 10;
    const chunks = [];
    for (let i = 0; i < problemList.length; i += CHUNK_SIZE) {
        chunks.push(problemList.slice(i, i + CHUNK_SIZE));
    }

    // Gemini 3.0 Flash Preview (User Specified)
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;

    const requests = chunks.map((chunk, i) => {
        // プロンプト構築
        let promptText = `
あなたは電気工学の専門家かつ厳格な採点者です。以下の試験問題に対する学生の回答を一括で採点してください。
各問題に対して、必ず JSON 配列の形式で [index, score, reason] を返してください。

【採点対象リスト】
`;
        chunk.forEach((p, index) => {
            const ans = typeof p.studentAnswer === 'string' ? p.studentAnswer : JSON.stringify(p.studentAnswer);
            promptText += `
---
ID: ${index}
[問題ID: ${p.qId}${p.sqId ? '_' + p.sqId : ''}]
問題文: ${p.text}
配点: ${p.points}点
採点基準: ${p.criteria}
学生の回答: ${ans || '(未回答)'}
`;
        });

        // リクエストペイロード (Strict JSON Mode)
        return {
            url: url,
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({
                contents: [{ parts: [{ text: promptText }] }],
                generationConfig: {
                    response_mime_type: "application/json",
                    response_schema: {
                        type: "ARRAY",
                        items: {
                            type: "OBJECT",
                            properties: {
                                index: { type: "INTEGER" },
                                score: { type: "INTEGER" },
                                reason: { type: "STRING" }
                            },
                            required: ["index", "score", "reason"]
                        }
                    }
                }
            }),
            muteHttpExceptions: true
        };
    });

    // リトライ付き並列実行
    let responses = [];
    try {
        responses = _fetchAllWithRetry(requests);
    } catch (e) {
        console.error("Critical Grading Error:", e);
        return problemList.map(p => ({
            questionId: p.qId,
            subQuestionId: p.sqId,
            score: 0,
            reason: "システムエラー: 採点処理に失敗しました。",
            questionText: p.text,
            modelAnswer: p.modelAnswer
        }));
    }

    let allResults = [];

    // レスポンス処理
    responses.forEach((response, i) => {
        const chunk = chunks[i];

        // レスポンスがない、またはエラーの場合
        if (!response || response.getResponseCode() !== 200) {
            console.error(`Chunk ${i} Failed:`, response ? response.getContentText() : "No response");
            allResults = allResults.concat(chunk.map(p => ({
                questionId: p.qId,
                subQuestionId: p.sqId,
                score: 0,
                reason: "採点エラー: AIからの応答がありませんでした。",
                questionText: p.text,
                modelAnswer: p.modelAnswer
            })));
            return;
        }

        try {
            const data = JSON.parse(response.getContentText());
            let parsedData = [];

            if (data.candidates && data.candidates[0].content && data.candidates[0].content.parts[0].text) {
                const text = data.candidates[0].content.parts[0].text;
                parsedData = JSON.parse(text);
            }

            if (!Array.isArray(parsedData)) {
                parsedData = [];
            }

            // マッピング
            const mapped = parsedData.map(r => {
                const idx = r.index;
                if (typeof idx !== 'number' || idx < 0 || idx >= chunk.length) return null;
                const original = chunk[idx];
                return {
                    questionId: original.qId,
                    subQuestionId: original.sqId,
                    score: Number(r.score) || 0,
                    reason: r.reason || '',
                    questionText: original.text,    // Include text
                    modelAnswer: original.modelAnswer // Include model answer
                };
            }).filter(Boolean);

            // マッピング漏れ補完
            if (mapped.length < chunk.length) {
                chunk.forEach(p => {
                    const exists = mapped.find(m => m.questionId === p.qId && m.subQuestionId === p.sqId);
                    if (!exists) {
                        mapped.push({
                            questionId: p.qId,
                            subQuestionId: p.sqId,
                            score: 0,
                            reason: "採点不能: AIが回答を認識できませんでした。",
                            questionText: p.text,
                            modelAnswer: p.modelAnswer
                        });
                    }
                });
            }

            allResults = allResults.concat(mapped);

        } catch (e) {
            console.error(`Chunk ${i} Parsing Error:`, e);
            allResults = allResults.concat(chunk.map(p => ({
                questionId: p.qId,
                subQuestionId: p.sqId,
                score: 0,
                reason: "採点エラー: 結果の解析に失敗しました。",
                questionText: p.text,
                modelAnswer: p.modelAnswer
            })));
        }
    });

    return allResults;
}

/**
 * UrlFetchApp.fetchAll ラッパー: 指数バックオフによるリトライ実装
 * @param {Array} requests - UrlFetchAppのリクエストオブジェクト配列
 * @param {number} maxRetries - 最大リトライ回数 (デフォルト3)
 * @return {Array} responses - UrlFetchApp.HTTPResponseの配列 (失敗時はnullが入る可能性あり)
 */
function _fetchAllWithRetry(requests, maxRetries = 3) {
    let results = new Array(requests.length).fill(null);
    let pendingIndices = requests.map((_, i) => i);

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        if (pendingIndices.length === 0) break;

        const currentRequests = pendingIndices.map(i => requests[i]);
        let batchResponses = [];
        let fetchError = null;

        try {
            // 並列リクエスト実行
            batchResponses = UrlFetchApp.fetchAll(currentRequests);
        } catch (e) {
            fetchError = e;
            console.warn(`FetchAll Attempt ${attempt + 1} Error: ${e}`);
        }

        const nextPending = [];

        if (fetchError) {
            // fetchAll自体が落ちた場合 (DNSエラー等)、全件リトライ対象
            // ただし、fetchAllは全か無かなので、個別のステータスコードは不明。
            // ネットワークエラーとみなして全件リトライ。
            nextPending.push(...pendingIndices);
        } else {
            // 個別のレスポンスをチェック
            batchResponses.forEach((res, i) => {
                const originalIndex = pendingIndices[i];
                const code = res.getResponseCode();

                // 429 (Too Many Requests) または 5xx (Server Error) はリトライ
                if (code === 429 || code >= 500) {
                    console.warn(`Request ${originalIndex} failed with ${code}. Retrying...`);
                    nextPending.push(originalIndex);
                } else {
                    results[originalIndex] = res;
                }
            });
        }

        pendingIndices = nextPending;

        if (pendingIndices.length > 0) {
            if (attempt < maxRetries) {
                // 指数バックオフ (1s, 2s, 4s...) + ジッター
                const sleepMs = Math.pow(2, attempt) * 1000 + Math.floor(Math.random() * 1000);
                Utilities.sleep(sleepMs);
            } else {
                console.error("Max retries reached. Some requests failed.");
            }
        }
    }

    return results;
}

/**
 * 回答ログをシートに保存
 */
function _saveResponseLog(questions, answers, gradingResults, totalScore, sessionId) {
    const ssId = _getSpreadsheetId();
    const ss = SpreadsheetApp.openById(ssId);
    let sheet = ss.getSheetByName(SHEET_NAME_RESPONSES);
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME_RESPONSES);
        sheet.appendRow(['Timestamp', 'Total Score', 'SessionID', 'Details (JSON)']); // Header
    }

    const detailObj = {
        answers: answers,
        grading: gradingResults
    };

    sheet.appendRow([
        new Date(),
        totalScore,
        sessionId || '',
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
    // Generate 15 dummy questions to test multi-chunk parallel processing (CHUNK_SIZE=7 -> 3 chunks)
    const dummyQuestions = Array.from({ length: 15 }, (_, i) => ({
        id: `debug_q${i + 1}`,
        text: `電流の単位は何か？(Test Q${i + 1})`,
        points: 10,
        criteria: '正解は「A」'
    }));

    const dummyAnswers = {};
    dummyQuestions.forEach(q => {
        dummyAnswers[q.id] = 'A';
    });

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

const SHEET_NAME_DIAGRAMS = 'ReferenceDiagrams';

/**
 * 参考図リストの保存
 * @param {Array} diagrams - [{id, name, url}, ...]
 */
function saveReferenceDiagrams(diagrams) {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        let sheet = ss.getSheetByName(SHEET_NAME_DIAGRAMS);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME_DIAGRAMS);
        }

        sheet.clear();
        sheet.appendRow(['ID', 'Name', 'URL']);

        if (diagrams && diagrams.length > 0) {
            const rows = diagrams.map(d => [d.id, d.name, d.url]);
            sheet.getRange(2, 1, rows.length, 3).setValues(rows);
        }

        SpreadsheetApp.flush();
        return { success: true, message: '図リストを保存しました。' };
    } catch (e) {
        console.error(e);
        return { success: false, message: '図リスト保存エラー: ' + e.toString() };
    }
}

/**
 * 参考図リストの取得
 */
function getReferenceDiagrams() {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        const sheet = ss.getSheetByName(SHEET_NAME_DIAGRAMS);
        if (!sheet) return [];

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
        return data.map(row => ({
            id: row[0],
            name: row[1],
            url: row[2]
        }));
    } catch (e) {
        console.error(e);
        return [];
    }
}
