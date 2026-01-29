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
const SHEET_NAME_DEBUG_LOG = 'DebugLog';

function _log(message) {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        let sheet = ss.getSheetByName(SHEET_NAME_DEBUG_LOG);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME_DEBUG_LOG);
            sheet.appendRow(['Timestamp', 'Message']);
        }
        sheet.appendRow([new Date(), message]);
    } catch (e) {
        console.error("Log failed", e);
    }
}

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
                q.subQuestions ? JSON.stringify(q.subQuestions) : '',
                q.modelAnswer || ''
            ]);
            sheet.getRange(2, 1, rows.length, 7).setValues(rows);
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
    _log(`submitAnswers started (Session: ${sessionId})`);
    try {
        const allQuestions = getQuestions();
        _log(`Loaded ${allQuestions.length} questions.`);

        // 【デバッグ用】 処理を5問だけに制限 -> 戻すことも可能だが一旦安全のため維持
        const questions = allQuestions; // 全件処理
        _log(`Processing all ${questions.length} questions.`);

        // 1. Gemini APIによる採点 (各問題ごと)
        _log("Calling _gradeWithGemini...");
        const gradingResults = _gradeWithGemini(questions, answers);
        _log(`_gradeWithGemini returned ${gradingResults.length} results.`);

        const totalScore = gradingResults.reduce((sum, r) => sum + r.score, 0);
        _log(`Total Score calculated: ${totalScore}`);

        // 2. 点数表の更新 (sessionIdがある場合)
        if (sessionId) {
            _updateScoreTable(sessionId, totalScore);
            _log("Score table updated.");
        }

        // 3. 詳細ログをスプレッドシートに保存 (バックアップ/詳細分析用)
        _saveResponseLog(questions, answers, gradingResults, totalScore, sessionId);
        _log("Response log saved. Returning success.");

        // GASの通信トラブル回避のため、JSON文字列として返す
        return JSON.stringify({
            success: true,
            totalScore: totalScore,
            results: gradingResults
        });

    } catch (e) {
        _log(`submitAnswers FATAL ERROR: ${e.toString()}`);
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
 * Gemini APIと通信して採点を行う内部関数 (一括採点・点数のみ版)
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
                    tempId: `${q.id}_${sq.id}`,
                    text: `[親問題]: ${q.text}\n[小問題]: ${sq.text}`,
                    points: sq.points,
                    criteria: sq.criteria || '特になし',
                    studentAnswer: (answers && answers[q.id] && answers[q.id][sq.id]) ? answers[q.id][sq.id] : ""
                });
            });
        } else {
            problemList.push({
                type: 'normal',
                qId: q.id,
                sqId: null,
                tempId: String(q.id),
                text: q.text,
                points: q.points,
                criteria: q.criteria || '特になし',
                studentAnswer: (answers && answers[q.id]) ? answers[q.id] : ""
            });
        }
    });

    if (problemList.length === 0) return [];

    // チャンク分割
    // 解説を生成しないため、トークン消費が少ない。1回に10問程度まとめて送る。
    const CHUNK_SIZE = 10;
    const chunks = [];
    for (let i = 0; i < problemList.length; i += CHUNK_SIZE) {
        chunks.push(problemList.slice(i, i + CHUNK_SIZE));
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

    console.log(`Starting Parallel Grading for ${chunks.length} chunks (fetchAll)...`);
    _log(`_gradeWithGemini: Prepared ${chunks.length} chunks. Creating requests...`);

    // リクエストの構築
    const requests = chunks.map((chunk, i) => {
        let promptText = `
你是電気工学の専門家かつ厳格な採点者です。以下の試験問題に対する学生の回答を一括で採点してください。

【重要：出力ルール】
- **点数のみ**を判定してください。
- 解説や理由は**一切不要**です。
- 出力はJSON配列のみとしてください。

【採点対象リスト】
`;
        chunk.forEach((p) => {
            const ans = typeof p.studentAnswer === 'string' ? p.studentAnswer : JSON.stringify(p.studentAnswer);
            promptText += `
---
[ID: ${p.tempId}]
問題文: ${p.text}
配点: ${p.points}点
採点基準: ${p.criteria}
学生の回答: ${ans || '(未回答)'}
`;
        });

        promptText += `
---

【JSON出力フォーマット】
[
  { "id": "ID文字列", "score": 数値 },
  { "id": "ID文字列", "score": 数値 },
  ...
]
※ "id" は上記リストの [ID: ...] と完全に一致させてください。
※ 未回答の場合は0点としてください。
`;

        const payload = JSON.stringify({
            contents: [{ parts: [{ text: promptText }] }],
            generationConfig: { response_mime_type: "application/json" }
        });

        return {
            url: url,
            method: 'post',
            contentType: 'application/json',
            payload: payload,
            muteHttpExceptions: true
        };
    });

    let allResults = [];
    let responses = [];

    // 並列リクエスト実行
    try {
        _log("Executing UrlFetchApp.fetchAll...");
        responses = UrlFetchApp.fetchAll(requests);
        _log(`fetchAll completed. Received ${responses.length} responses.`);
    } catch (e) {
        _log(`fetchAll EXCEPTION: ${e.toString()}`);
        console.error("fetchAll failed:", e);
        // 全失敗として扱う
        return problemList.map(p => ({
            questionId: p.qId,
            subQuestionId: p.sqId,
            score: 0,
            reason: `判定エラー (通信一括失敗: ${e.message})`
        }));
    }

    // レスポンス処理
    responses.forEach((response, i) => {
        const chunk = chunks[i];
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        let chunkResults = [];
        let success = false;

        if (responseCode === 200) {
            try {
                const data = JSON.parse(responseText);
                if (data.candidates && data.candidates.length > 0) {
                    const text = data.candidates[0].content.parts[0].text;
                    const jsonStr = text.replace(/```json/g, '').replace(/```/g, '').trim();

                    let parsedData;
                    try {
                        parsedData = JSON.parse(jsonStr);
                    } catch (e) {
                        const match = jsonStr.match(/\[[\s\S]*\]/);
                        if (match) parsedData = JSON.parse(match[0].replace(/\\/g, '\\\\'));
                    }

                    if (Array.isArray(parsedData)) {
                        chunkResults = parsedData;
                        success = true;
                    } else if (parsedData) {
                        chunkResults = [parsedData];
                        success = true;
                    }
                }
            } catch (parseError) {
                console.error(`Chunk ${i} parse error:`, parseError);
            }
        } else {
            console.error(`Chunk ${i} error: ${responseCode} - ${responseText}`);
        }

        if (success) {
            // IDベースでマッピング
            const mapped = chunk.map(original => {
                const match = chunkResults.find(r => String(r.id) === String(original.tempId));
                let score = 0;
                let reason = "";

                if (match) {
                    score = Number(match.score) || 0;
                    if (score === original.points) reason = "正解 (AI判定)";
                    else if (score === 0) reason = "不正解 (AI判定)";
                    else reason = "部分点 (AI判定)";
                } else {
                    console.warn(`Missing result for ID: ${original.tempId}`);
                    reason = "判定エラー (結果なし)";
                }

                return {
                    questionId: original.qId,
                    subQuestionId: original.sqId,
                    score: score,
                    reason: reason
                };
            });
            allResults = allResults.concat(mapped);
        } else {
            const fallback = chunk.map(p => ({
                questionId: p.qId,
                subQuestionId: p.sqId,
                score: 0,
                reason: `判定エラー (ステータス: ${responseCode})`
            }));
            allResults = allResults.concat(fallback);
        }
    });

    return allResults;
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
 * ポーリング用: 採点結果の確認
 * @param {string} sessionId
 */
function checkGradingStatus(sessionId) {
    try {
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);
        const sheet = ss.getSheetByName(SHEET_NAME_RESPONSES);

        if (!sheet) return JSON.stringify({ status: "PENDING" });

        // 後ろから検索（最新のものが該当するはず）
        const lastRow = sheet.getLastRow();
        const data = sheet.getRange(Math.max(2, lastRow - 20), 1, Math.min(21, lastRow - 1), 4).getValues();

        for (let i = data.length - 1; i >= 0; i--) {
            const row = data[i];
            // Col 3 is SessionID
            if (String(row[2]) === String(sessionId)) {
                // Col 2 (index 1) is Total Score. If "PENDING", still processing.
                const scoreCell = row[1];

                if (scoreCell === "PENDING") {
                    return JSON.stringify({ status: "PROCESSING" });
                }

                // 採点完了していれば詳細JSONを返す
                // Col 4 (index 3) is JSON details
                try {
                    const detail = JSON.parse(row[3]);
                    if (detail.grading) {
                        return JSON.stringify({
                            status: "COMPLETED",
                            totalScore: scoreCell,
                            results: detail.grading
                        });
                    }
                } catch (e) {
                    console.error("JSON parse error in checkStatus", e);
                }
            }
        }

        return JSON.stringify({ status: "NOT_FOUND" });

    } catch (e) {
        return JSON.stringify({ status: "ERROR", message: e.toString() });
    }
}

/**
 * デバッグ用: 採点ロジック単体テスト
 */
function testGrading() {
    // Generate 60 dummy questions to force multiple chunks (CHUNK_SIZE=10 -> 6 chunks)
    // This stress tests the parallelism and API rate limits
    const dummyQuestions = Array.from({ length: 60 }, (_, i) => ({
        id: `debug_q${i + 1}`,
        text: `電気回路におけるオームの法則について説明し、電圧V=10V, 抵抗R=5Ωの時の電流Iを求めよ。(Stress Test Q${i + 1})`,
        points: 10,
        criteria: 'オームの法則(V=IR)への言及があること。計算結果が2Aであること。'
    }));

    const dummyAnswers = {};
    dummyQuestions.forEach(q => {
        dummyAnswers[q.id] = 'オームの法則は電圧と電流と抵抗の関係を示す。I = V/R = 10/5 = 2A です。';
    });

    try {
        const start = new Date();
        console.log("Starting Stress Test with 60 questions...");
        // testGradingの中で submitAnswers を呼んでみる
        const results = submitAnswers(dummyAnswers);
        const end = new Date();
        const duration = (end - start) / 1000;

        console.log(`Stress Test Completed in ${duration}s`);
        return {
            success: true,
            message: `負荷テスト(60問) 成功 (${duration}秒)`,
            details: results
        };
    } catch (e) {
        return {
            success: false,
            message: "負荷テスト失敗: " + e.toString()
        };
    }
}

/**
 * システム診断関数
 * 本番データ（Spreadsheet）の読み込み状況などをチェックします。
 */
function diagnoseSystem() {
    try {
        const start = new Date();
        const ssId = _getSpreadsheetId();
        const ss = SpreadsheetApp.openById(ssId);

        // 1. Check Questions Sheet
        const qSheet = ss.getSheetByName(SHEET_NAME_QUESTIONS);
        if (!qSheet) return "Error: Question sheet not found";

        const lastRow = qSheet.getLastRow();
        const questionCount = lastRow - 1; // Exclude header

        // Load Questions
        const questions = getQuestions();
        const loadedCount = questions.length;

        // Check for anomalies
        let subQuestionTotal = 0;
        let maxTextLength = 0;
        questions.forEach(q => {
            if (q.subQuestions) subQuestionTotal += q.subQuestions.length;
            if (q.text.length > maxTextLength) maxTextLength = q.text.length;
        });

        // 2. Check API Key
        const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEY_GEMINI_API_KEY);
        const hasKey = !!apiKey;

        const end = new Date();

        const result = {
            success: true,
            checkTime: (end - start) / 1000 + "s",
            sheetId: ssId,
            sheetRows: lastRow,
            loadedQuestions: loadedCount,
            totalSubQuestions: subQuestionTotal,
            maxTextLength: maxTextLength,
            hasApiKey: hasKey,
            firstQuestion: questions.length > 0 ? questions[0].text.substring(0, 50) + "..." : "None"
        };
        console.log(JSON.stringify(result, null, 2));
        return result;

    } catch (e) {
        const errorResult = {
            success: false,
            message: "Diagnosis Failed: " + e.toString(),
            stack: e.stack
        };
        console.error(JSON.stringify(errorResult));
        return errorResult;
    }
}
