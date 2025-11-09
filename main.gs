// =================================================
// Global Settings
// =================================================

// 状態を保存するためのキー
const STATE_KEY = 'SYNC_STATE';

// 処理を再開するためのトリガーとして設定する関数名
const TRIGGER_HANDLER_NAME = 'mainTriggerHandler';

// 実行時間の上限（ミリ秒）。GASの上限は約6分のため、安全マージンをとって5分に設定。
const EXECUTION_LIMIT_MS = 5 * 60 * 1000;


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
    // Overall Status
    status: 'IDLE', // IDLE, RUNNING, DONE, ERROR
    phase: 'NONE',   // NONE, PLANNING, EXECUTING
    message: '待機中',
    lastError: null,
    startTime: null,

    // Folder Info
    sourceFolderId: null,
    destFolderId: null,

    // Planning Phase
    sourceMap: {},
    destMap: {},
    scanQueue: [],

    // Executing Phase
    actions: [],
    processedActions: 0,

    // Stats
    totalFolders: 0,
    totalFiles: 0,
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

    const sourceRoot = Drive.Files.get(sourceFolderId, { supportsAllDrives: true, fields: 'id, name' });
    const destRoot = Drive.Files.get(destFolderId, { supportsAllDrives: true, fields: 'id, name' });

    const initialState = {
      ...getDefaultState(),
      status: 'RUNNING',
      phase: 'PLANNING',
      message: '同期計画を準備中です...',
      startTime: Date.now(),
      sourceFolderId: sourceFolderId,
      destFolderId: destFolderId,
      scanQueue: [
        { type: 'source', folderId: sourceFolderId, path: '' },
        { type: 'dest', folderId: destFolderId, path: '' },
      ],
      sourceMap: { '': {id: sourceFolderId, name: sourceRoot.name, type: 'folder', children: {}} },
      destMap: { '': {id: destFolderId, name: destRoot.name, type: 'folder', children: {}} },
    };
    setState(initialState);

    ScriptApp.newTrigger(TRIGGER_HANDLER_NAME).timeBased().after(1000).create();
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

function mainTriggerHandler() {
  const startTime = Date.now();
  let state = getState();

  if (state.status !== 'RUNNING') {
    Logger.log(`Execution skipped. Status is "${state.status}", not "RUNNING".`);
    deleteTriggers();
    return;
  }

  try {
    if (state.phase === 'PLANNING') {
      state = runPlanningPhase(state, startTime);
    }

    if (state.phase === 'GENERATING_ACTIONS') {
      state = runGenerateActionsPhase(state);
    }

    if (state.phase === 'EXECUTING') {
      state = runExecutingPhase(state, startTime);
    }

    if (state.status === 'RUNNING') {
      deleteTriggers();
      state.message = "一時中断中... 次の処理を準備しています。";
      ScriptApp.newTrigger(TRIGGER_HANDLER_NAME).timeBased().after(1000).create();
    } else {
      deleteTriggers();
    }

  } catch (e) {
    Logger.log(`ERROR in mainTriggerHandler: ${e.stack}`);
    state.status = 'ERROR';
    state.message = `エラーが発生しました: ${e.message}`;
    state.lastError = e.stack;
    deleteTriggers();
  } finally {
    setState(state);
  }
}

function runPlanningPhase(state, startTime) {
    while (state.scanQueue.length > 0 && (Date.now() - startTime) < EXECUTION_LIMIT_MS) {
        const item = state.scanQueue.shift();
        const parentMap = (item.type === 'source') ? state.sourceMap : state.destMap;
        state.message = `フォルダ情報をスキャン中: ${item.path || '/'}`;

        const children = getChildren(item.folderId);
        let currentPathMap = getObjectByPath(parentMap, item.path);

        for (const child of children) {
            const childPath = item.path ? `${item.path}/${child.name}` : child.name;
            const isFolder = child.mimeType === 'application/vnd.google-apps.folder';
            currentPathMap.children[child.name] = { id: child.id, type: isFolder ? 'folder' : 'file', modifiedTime: child.modifiedTime, children: isFolder ? {} : undefined };
            if (isFolder) {
                state.scanQueue.push({ type: item.type, folderId: child.id, path: childPath });
                if (item.type === 'source') state.totalFolders++;
            } else {
                if (item.type === 'source') state.totalFiles++;
            }
        }
    }
    if (state.scanQueue.length === 0) {
        state.phase = 'GENERATING_ACTIONS';
    }
    return state;
}

function runGenerateActionsPhase(state) {
    state.message = "変更点を分析し、同期計画を作成しています...";
    const actions = [];
    const addActionsForNewFolder = (sourceNode, parentPath) => {
        for (const name in sourceNode.children) {
            const child = sourceNode.children[name];
            const currentPath = parentPath ? `${parentPath}/${name}` : name;
            if (child.type === 'folder') {
                actions.push({ type: 'CREATE_FOLDER', path: currentPath });
                addActionsForNewFolder(child, currentPath);
            } else {
                actions.push({ type: 'COPY_FILE', path: currentPath, sourceId: child.id });
            }
        }
    };
    const compareFolders = (sourceNode, destNode, path) => {
        for (const name in sourceNode.children) {
            const sourceChild = sourceNode.children[name];
            const destChild = destNode.children[name];
            const currentPath = path ? `${path}/${name}` : name;
            if (!destChild) {
                if (sourceChild.type === 'folder') {
                    actions.push({ type: 'CREATE_FOLDER', path: currentPath });
                    addActionsForNewFolder(sourceChild, currentPath);
                } else {
                    actions.push({ type: 'COPY_FILE', path: currentPath, sourceId: sourceChild.id });
                }
            } else {
                 if (sourceChild.type === 'file' && destChild.type === 'file') {
                    const sourceModified = new Date(sourceChild.modifiedTime).getTime();
                    const destModified = new Date(destChild.modifiedTime).getTime();
                    if (sourceModified > destModified) {
                        actions.push({ type: 'UPDATE_FILE', path: currentPath, sourceId: sourceChild.id, destId: destChild.id });
                    }
                } else if (sourceChild.type === 'folder' && destChild.type === 'folder') {
                    compareFolders(sourceChild, destChild, currentPath);
                }
            }
        }
    };
    compareFolders(state.sourceMap[''], state.destMap[''], '');
    state.actions = actions;
    state.phase = 'EXECUTING';
    return state;
}

function runExecutingPhase(state, startTime) {
    try {
        while (state.actions.length > state.processedActions && (Date.now() - startTime) < EXECUTION_LIMIT_MS) {
            const action = state.actions[state.processedActions];
            const progress = `(${state.processedActions + 1}/${state.actions.length})`;

            const parentPath = getParentPath(action.path);
            const fileName = getFileName(action.path);
            const parentNode = getObjectByPath(state.destMap, parentPath);

            if (!parentNode) {
                throw new Error(`親フォルダが見つかりません: ${action.path}`);
            }

            state.message = `処理中: ${action.path} ${progress}`;

            switch(action.type) {
                case 'CREATE_FOLDER':
                    const newFolder = { name: fileName, parents: [parentNode.id], mimeType: 'application/vnd.google-apps.folder' };
                    const createdFolder = Drive.Files.create(newFolder, null, { supportsAllDrives: true, fields: 'id' });
                    parentNode.children[fileName] = { id: createdFolder.id, type: 'folder', children: {} };
                    break;

                case 'COPY_FILE':
                    const copyResource = { name: fileName, parents: [parentNode.id] };
                    Drive.Files.copy(copyResource, action.sourceId, { supportsAllDrives: true });
                    break;

                case 'UPDATE_FILE':
                     Drive.Files.remove(action.destId, { supportsAllDrives: true });
                     const updateResource = { name: fileName, parents: [parentNode.id] };
                     Drive.Files.copy(updateResource, action.sourceId, { supportsAllDrives: true });
                    break;
            }
            state.processedActions++;
        }
    } catch (e) {
        const failedAction = state.actions[state.processedActions];
        Logger.log(`実行エラー: アクション=${JSON.stringify(failedAction)}, エラー=${e.stack}`);
        throw new Error(`アクション ${failedAction.type} に失敗しました: ${failedAction.path}. 原因: ${e.message}`);
    }

    if (state.actions.length <= state.processedActions) {
        state.status = 'DONE';
        const elapsedTime = (Date.now() - state.startTime) / 1000;
        state.message = `同期が完了しました。(${state.totalFolders}フォルダ, ${state.totalFiles}ファイル) 経過時間: ${Math.round(elapsedTime)}秒`;
    }
    return state;
}

// =================================================
// Helper Functions
// =================================================

function getChildren(folderId) {
  const children = [];
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
      children.push(...response.files);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  return children;
}

function getObjectByPath(obj, path) {
    if (path === '') return obj[''];
    return path.split('/').reduce((acc, part) => {
        if (!acc || !acc.children || !acc.children[part]) return null;
        return acc.children[part];
    }, obj['']);
}

function getParentPath(path) {
    const parts = path.split('/');
    parts.pop();
    return parts.join('/');
}

function getFileName(path) {
    return path.split('/').pop();
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_HANDLER_NAME) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}
