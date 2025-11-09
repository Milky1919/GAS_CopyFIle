// =================================================
// Global Settings
// =================================================

// 状態を保存するためのキー
const STATE_KEY = 'SYNC_STATE';

// 処理を再開するためのトリガーとして設定する関数名
const TRIGGER_HANDLER_NAME = 'processFolderQueue';

// 実行時間の上限（ミリ秒）。GASの上限30分に対し、安全マージンをとって25分に設定。
const EXECUTION_LIMIT_MS = 25 * 60 * 1000;


// =================================================
// UI Control Functions
// =================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('フォルダ同期').addItem('同期を開始', 'showDialog').addToUi();
}

function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('dialog').setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'フォルダ同期設定');
}


// =================================================
// State Management Functions
// =================================================

function getDefaultState() {
  return {
    status: 'IDLE', message: '待機中',
    sourceFolderId: null, destFolderId: null,
    folderQueue: [], processedFolders: 0,
    processedFiles: 0, lastError: null,
  };
}

function getState() {
  const properties = PropertiesService.getScriptProperties();
  const stateJson = properties.getProperty(STATE_KEY);
  return stateJson ? JSON.parse(stateJson) : getDefaultState();
}

function setState(state) {
  const stateJson = JSON.stringify(state, null, 2);
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, stateJson);
  Logger.log(`STATE UPDATED: ${state.status} - ${state.message}`);
}


// =================================================
// Functions called from HTML Dialog
// =================================================

function getFolderIdFromInput(input) {
  input = input.trim();
  if (input.includes('folders/')) return input.split('folders/')[1].split('/')[0].split('?')[0];
  if (input.includes('id=')) return input.split('id=')[1].split('&')[0];
  return input;
}

function getSyncStatus() { return getState(); }

function startSyncJob(folderIds) {
  Logger.log(`startSyncJob called with: ${JSON.stringify(folderIds)}`);
  try {
    deleteTriggers();
    const sourceFolderId = getFolderIdFromInput(folderIds.sourceFolder);
    const destFolderId = getFolderIdFromInput(folderIds.destFolder);

    try { Drive.Files.get(sourceFolderId, { supportsAllDrives: true, fields: 'id' }); }
    catch (e) { throw new Error(`コピー元フォルダ(ID: ${sourceFolderId})にアクセスできません。`); }
    try { Drive.Files.get(destFolderId, { supportsAllDrives: true, fields: 'id' }); }
    catch (e) { throw new Error(`同期先フォルダ(ID: ${destFolderId})にアクセスできません。`); }

    const initialState = {
      ...getDefaultState(), status: 'RUNNING', message: '同期の準備をしています...',
      sourceFolderId: sourceFolderId, destFolderId: destFolderId,
      folderQueue: [{ sourceId: sourceFolderId, destId: destFolderId, path: '' }],
    };
    setState(initialState);

    // 最初のトリガーを10秒後に設定
    const trigger = ScriptApp.newTrigger(TRIGGER_HANDLER_NAME).timeBased().after(10 * 1000).create();
    Logger.log(`Trigger created successfully. ID: ${trigger.getUniqueId()}`);
  } catch (e) {
    Logger.log(`Error in startSyncJob: ${e.stack}`);
    setState({ ...getState(), status: 'ERROR', message: e.message, lastError: e.message });
    throw e;
  }
}

function stopSyncJob() {
  deleteTriggers();
  setState(getDefaultState());
  Logger.log('同期ジョブがユーザーによって停止されました。');
}


// =================================================
// Core Sync Logic (Triggered Execution)
// =================================================

/**
 * フォルダキューを処理するメイン関数。トリガーによって定期的に実行される。
 */
function processFolderQueue() {
  Logger.log('--- processFolderQueue triggered ---');
  const startTime = Date.now();
  let state = getState();

  Logger.log(`Current state: ${JSON.stringify(state, null, 2)}`);

  if (state.status !== 'RUNNING') {
    Logger.log(`Execution skipped. Status is "${state.status}", not "RUNNING".`);
    deleteTriggers(); // Clean up unnecessary triggers.
    return;
  }

  try {
    // フォルダキューが空になるか、実行時間制限に達するまでループ
    while (state.folderQueue.length > 0 && (Date.now() - startTime) < EXECUTION_LIMIT_MS) {
      const currentFolder = state.folderQueue[0]; // キューの先頭を取得

      // 宛先フォルダ内の既存アイテムのマップを作成
      const destChildrenMap = getChildrenMap(currentFolder.destId);

      let pageToken = null;
      do {
        const listParams = {
          q: `'${currentFolder.sourceId}' in parents and trashed = false`,
          fields: 'nextPageToken, files(id, name, mimeType, modifiedTime)',
          supportsAllDrives: true, includeItemsFromAllDrives: true,
          pageSize: 200, pageToken: pageToken,
        };
        const response = Drive.Files.list(listParams);

        if (response.files) {
          for (const sourceFile of response.files) {
            const destFile = destChildrenMap.get(sourceFile.name);
            const currentPath = `${currentFolder.path}/${sourceFile.name}`;
            state.message = `処理中: ${currentPath}`;

            if (sourceFile.mimeType === 'application/vnd.google-apps.folder') {
              // --- フォルダの処理 ---
              let destSubFolderId;
              if (destFile && destFile.mimeType === 'application/vnd.google-apps.folder') {
                destSubFolderId = destFile.id;
              } else {
                // 宛先に同名ファイルがある場合は削除（またはリネーム）も考慮できるが、今回はシンプルにフォルダ作成
                if(destFile) Drive.Files.remove(destFile.id, { supportsAllDrives: true });
                const newFolder = { name: sourceFile.name, parents: [currentFolder.destId], mimeType: 'application/vnd.google-apps.folder' };
                destSubFolderId = Drive.Files.create(newFolder, null, { supportsAllDrives: true, fields: 'id' }).id;
              }
              state.folderQueue.push({ sourceId: sourceFile.id, destId: destSubFolderId, path: currentPath });
            } else {
              // --- ファイルの処理 ---
              let shouldCopy = false;
              if (!destFile) {
                shouldCopy = true; // 存在しない
              } else if (new Date(sourceFile.modifiedTime).getTime() > new Date(destFile.modifiedTime).getTime()) {
                Drive.Files.remove(destFile.id, { supportsAllDrives: true }); // 古いファイルを削除
                shouldCopy = true; // 更新日時が新しい
              }

              if (shouldCopy) {
                const copyResource = { name: sourceFile.name, parents: [currentFolder.destId] };
                Drive.Files.copy(copyResource, sourceFile.id, { supportsAllDrives: true });
              }
              state.processedFiles++;
            }
          }
        }
        pageToken = response.nextPageToken;
        setState(state); // 1ページ処理ごとに状態を保存
      } while (pageToken && (Date.now() - startTime) < EXECUTION_LIMIT_MS);

      // 処理が完了したフォルダをキューから削除
      if (!pageToken) { // ページネーションが完了した場合のみ
        state.folderQueue.shift();
        state.processedFolders++;
      }
      setState(state); // フォルダ処理完了/中断時に状態を保存
    }

    // ループ終了後の処理
    if (state.folderQueue.length > 0) {
      // まだキューに残っている -> 次のトリガーを設定
      deleteTriggers();
      ScriptApp.newTrigger(TRIGGER_HANDLER_NAME).timeBased().after(5 * 60 * 1000).create(); // 5分後
      state.message = "一時中断中... 5分後に自動で再開します。";
    } else {
      // キューが空になった -> 完了
      state.status = 'DONE';
      state.message = `同期が完了しました。(${state.processedFolders}フォルダ, ${state.processedFiles}ファイル)`;
      deleteTriggers();
    }

  } catch (e) {
    Logger.log(`ERROR in processFolderQueue: ${JSON.stringify(e, null, 2)}`);
    state.status = 'ERROR';
    state.message = `エラーが発生しました: ${e.message}`;
    state.lastError = e.stack;
    deleteTriggers();
  } finally {
    setState(state); // 最終的な状態を保存
  }
}

/**
 * 指定されたフォルダID直下の子要素のMapを返すヘルパー関数
 * @param {string} folderId
 * @returns {Map<string, object>} key: name, value: {id, modifiedTime, mimeType}
 */
function getChildrenMap(folderId) {
  const map = new Map();
  let pageToken = null;
  do {
    const listParams = {
      q: `'${folderId}' in parents and trashed = false`,
      fields: 'nextPageToken, files(id, name, mimeType, modifiedTime)',
      supportsAllDrives: true, includeItemsFromAllDrives: true,
      pageSize: 1000, pageToken: pageToken,
    };
    const response = Drive.Files.list(listParams);
    if (response.files) {
      for (const file of response.files) {
        map.set(file.name, { id: file.id, modifiedTime: file.modifiedTime, mimeType: file.mimeType });
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  return map;
}

// =================================================
// Trigger Management
// =================================================

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_HANDLER_NAME) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}
