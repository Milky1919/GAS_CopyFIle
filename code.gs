// ----------------------------------------
// メインのスクリプトファイル（コード.gs）
// (Drive API v3 専用版 - 差分更新機能追加)
// ----------------------------------------

/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('フォルダコピー (v3)')
    .addItem('[1] 新規コピーを実行', 'showCopyDialog')
    .addItem('[2] 差分更新を実行', 'showSyncDialog') // ★ 追加
    .addToUi();
}

// ===============================================
// [1] 新規コピー機能 (既存のコード)
// ===============================================

/**
 * (新規コピー用) HTMLのダイアログを表示します。
 */
function showCopyDialog() {
  const html = HtmlService.createHtmlOutputFromFile('dialog')
    .setWidth(450)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, ' [1] 新規コピー設定 (v3)');
}

/**
 * URLやIDの文字列からフォルダIDを抽出するヘルパー関数
 */
function getFolderIdFromInput(input) {
  // ... (既存のコード: 変更なし) ...
  input = input.trim();
  if (input.includes('folders/')) {
    return input.split('folders/')[1].split('/')[0].split('?')[0];
  } else if (input.includes('id=')) {
    return input.split('id=')[1].split('&')[0];
  }
  return input;
}

/**
 * (新規コピー用) メインのコピー処理関数 (Drive API v3版)
 */
function startFolderCopy(formData) {
  // ... (既存のコード: 変更なし) ...
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. 入力値の検証
    if (!formData.sourceFolder || !formData.destParentFolder || !formData.newFolderName) {
      throw new Error('すべてのフィールドを入力してください。');
    }

    // 2. フォルダIDの抽出
    const sourceFolderId = getFolderIdFromInput(formData.sourceFolder);
    const destParentFolderId = getFolderIdFromInput(formData.destParentFolder);
    const newFolderName = formData.newFolderName.trim();

    // 3. 権限チェック
    try { Drive.Files.get(sourceFolderId, { supportsAllDrives: true, fields: 'id' }); } catch (e) {
      throw new Error(`コピー元フォルダ(ID: ${sourceFolderId})にアクセスできません。`);
    }
    try { Drive.Files.get(destParentFolderId, { supportsAllDrives: true, fields: 'id' }); } catch (e) {
      throw new Error(`コピー先親フォルダ(ID: ${destParentFolderId})にアクセスできません。`);
    }

    // 4. コピー先のルートフォルダ作成
    let newRootFolder;
    try {
      const folderResource = {
        name: newFolderName,
        parents: [destParentFolderId],
        mimeType: 'application/vnd.google-apps.folder'
      };
      newRootFolder = Drive.Files.create(folderResource, null, {
        supportsAllDrives: true, 
        fields: 'id, webViewLink' 
      });
    } catch (e) {
      throw new Error(`コピー先親フォルダに「${newFolderName}」フォルダを作成できませんでした。 : ${e.message}`);
    }

    // 5. 再帰的にフォルダとファイルをコピー
    copyFolderRecursively_v3(sourceFolderId, newRootFolder.id);

    // 6. 完了通知
    const url = newRootFolder.webViewLink;
    const message = `コピーが完了しました。\n新しいフォルダはこちらです:\n${url}`;
    ui.alert('コピー完了', message, ui.ButtonSet.OK);
    
    return "コピーが完了しました。";

  } catch (e) {
    ui.alert(`エラーが発生しました`, e.message, ui.ButtonSet.OK);
    return e.message;
  }
}

/**
 * (新規コピー用) フォルダを再帰的にコピーする関数 (Drive API v3版)
 */
function copyFolderRecursively_v3(sourceFolderId, destParentFolderId) {
  // ... (既存のコード: 変更なし) ...
  let pageToken = null;
  do {
    const listParams = {
      q: `'${sourceFolderId}' in parents and trashed = false`,
      fields: 'nextPageToken, files(id, name, mimeType)',
      supportsAllDrives: true, includeItemsFromAllDrives: true,
      pageSize: 500, pageToken: pageToken
    };
    const response = Drive.Files.list(listParams);
    if (!response.files) { continue; }

    for (const file of response.files) {
      if (file.mimeType === 'application/vnd.google-apps.folder') {
        const subFolderResource = { name: file.name, parents: [destParentFolderId], mimeType: 'application/vnd.google-apps.folder' };
        const newSubFolder = Drive.Files.create(subFolderResource, null, { supportsAllDrives: true, fields: 'id' });
        copyFolderRecursively_v3(file.id, newSubFolder.id);
      } else {
        const copyResource = { name: file.name, parents: [destParentFolderId] };
        Drive.Files.copy(copyResource, file.id, { supportsAllDrives: true });
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
}


// ===============================================
// [2] 差分更新機能 (★ ここから新規追加 ★)
// ===============================================

// スクリプトプロパティ（進捗管理用）をリセットする
function resetProgress() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'sync_totalFiles': '0',
    'sync_processedFiles': '0',
    'sync_status': 'initializing'
  });
}

/**
 * (差分更新用) HTMLのダイアログを表示します。
 */
function showSyncDialog() {
  resetProgress(); // 実行前に必ずリセット
  const html = HtmlService.createHtmlOutputFromFile('dialog_sync')
    .setWidth(450)
    .setHeight(450); // 少し高さを増やす
  SpreadsheetApp.getUi().showModalDialog(html, ' [2] 差分更新 (v3)');
}

/**
 * (差分更新用) HTMLから進捗状況を取得するために呼び出されます。
 * @returns {object} 進捗状況 { total, processed, status }
 */
function getSyncProgress() {
  const properties = PropertiesService.getScriptProperties();
  return {
    total: parseInt(properties.getProperty('sync_totalFiles') || '0', 10),
    processed: parseInt(properties.getProperty('sync_processedFiles') || '0', 10),
    status: properties.getProperty('sync_status') || '...'
  };
}

/**
 * (差分更新用) コピー元の総ファイル数を事前にカウントします。
 * @param {string} folderId - カウント対象のフォルダID
 * @returns {number} 総ファイル数
 */
function getTotalFileCount(folderId) {
  let count = 0;
  let pageToken = null;
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('sync_status', 'コピー元の総ファイル数をカウント中...');

  do {
    const listParams = {
      q: `'${folderId}' in parents and trashed = false`,
      fields: 'nextPageToken, files(id, mimeType)',
      supportsAllDrives: true, includeItemsFromAllDrives: true,
      pageSize: 1000, pageToken: pageToken
    };
    const response = Drive.Files.list(listParams);
    if (!response.files) { continue; }

    for (const file of response.files) {
      if (file.mimeType === 'application/vnd.google-apps.folder') {
        count += getTotalFileCount(file.id); // 再帰的にカウント
      } else {
        count++; // ファイルをカウント
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return count;
}

/**
 * (差分更新用) メインの差分更新処理関数
 * @param {object} formData - { sourceFolder, destFolder }
 * @returns {string} 成功またはエラーメッセージ
 */
function startFolderSync(formData) {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  
  try {
    // 1. 入力値の検証
    if (!formData.sourceFolder || !formData.destFolder) {
      throw new Error('コピー元と更新先の両方のIDを入力してください。');
    }

    // 2. フォルダIDの抽出
    const sourceFolderId = getFolderIdFromInput(formData.sourceFolder);
    const destFolderId = getFolderIdFromInput(formData.destFolder); // 更新先はコピーされたルートフォルダ

    // 3. 権限チェック
    try { Drive.Files.get(sourceFolderId, { supportsAllDrives: true, fields: 'id' }); } catch (e) {
      throw new Error(`コピー元フォルダ(ID: ${sourceFolderId})にアクセスできません。`);
    }
    try { Drive.Files.get(destFolderId, { supportsAllDrives: true, fields: 'id' }); } catch (e) {
      throw new Error(`更新先フォルダ(ID: ${destFolderId})にアクセスできません。`);
    }
    
    // 4. (進捗用) コピー元の総ファイル数をカウント
    const totalFiles = getTotalFileCount(sourceFolderId);
    properties.setProperty('sync_totalFiles', totalFiles.toString());
    if (totalFiles === 0) {
      ui.alert('完了', 'コピー元にファイルが見つかりませんでした。', ui.ButtonSet.OK);
      return "完了 (ファイル0)";
    }

    // 5. (ステップ1) 更新先フォルダの全ファイル台帳を作成
    properties.setProperty('sync_status', '更新先のファイル一覧を作成中...');
    const destFileMap = new Map(); // キー: 相対パス, 値: { id, modifiedTime }
    buildDestMap_v3(destFolderId, destFileMap, '');

    // 6. (ステップ2) コピー元をスキャンし、差分を比較・実行
    properties.setProperty('sync_status', 'コピー元と比較し、差分更新を実行中...');
    syncRecursively_v3(sourceFolderId, destFolderId, destFileMap, '');

    // 7. 完了通知
    properties.setProperty('sync_status', '完了');
    const url = Drive.Files.get(destFolderId, { supportsAllDrives: true, fields: 'webViewLink' }).webViewLink;
    ui.alert('差分更新 完了', `更新が完了しました。\n更新先フォルダ:\n${url}`, ui.ButtonSet.OK);
    
    return "差分更新が完了しました。";

  } catch (e) {
    properties.setProperty('sync_status', `エラー: ${e.message}`);
    ui.alert(`エラーが発生しました`, e.message, ui.ButtonSet.OK);
    return e.message;
  }
}

/**
 * (差分更新用) 更新先のファイル台帳(Map)を作成する
 */
function buildDestMap_v3(folderId, map, currentPath) {
  let pageToken = null;
  do {
    const listParams = {
      q: `'${folderId}' in parents and trashed = false`,
      fields: 'nextPageToken, files(id, name, mimeType, modifiedTime)',
      supportsAllDrives: true, includeItemsFromAllDrives: true,
      pageSize: 1000, pageToken: pageToken
    };
    const response = Drive.Files.list(listParams);
    if (!response.files) { continue; }

    for (const file of response.files) {
      const path = `${currentPath}/${file.name}`;
      if (file.mimeType === 'application/vnd.google-apps.folder') {
        buildDestMap_v3(file.id, map, path); // 再帰
      } else {
        map.set(path, { id: file.id, modifiedTime: file.modifiedTime });
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
}

/**
 * (差分更新用) コピー元をスキャンし、台帳と比較しながら再帰的にコピー/更新する
 */
function syncRecursively_v3(sourceFolderId, destParentFolderId, destFileMap, currentPath) {
  const properties = PropertiesService.getScriptProperties();
  let pageToken = null;

  do {
    const listParams = {
      q: `'${sourceFolderId}' in parents and trashed = false`,
      fields: 'nextPageToken, files(id, name, mimeType, modifiedTime)',
      supportsAllDrives: true, includeItemsFromAllDrives: true,
      pageSize: 500, pageToken: pageToken
    };
    const response = Drive.Files.list(listParams);
    if (!response.files) { continue; }

    for (const file of response.files) {
      const path = `${currentPath}/${file.name}`;
      
      if (file.mimeType === 'application/vnd.google-apps.folder') {
        // --- フォルダの場合 ---
        properties.setProperty('sync_status', `フォルダ確認中: ${path}`);
        
        let newDestFolderId;
        const destFolderInfo = destFileMap.get(path); // フォルダも台帳でチェック (簡易的)
        
        // 簡易チェック: Mapはファイルのみだが、フォルダパスでファイルが存在すればフォルダもあるはず
        // ※厳密にはフォルダだけの存在チェックが必要だが、処理速度優先でファイルベースで判断
        // → やはりフォルダをちゃんとチェックする
        
        // フォルダの存在をAPIで確認 (これが確実だが遅い)
        // → やはり台帳作成時にフォルダも入れるべき
        
        // (仕様変更) buildDestMap_v3 もフォルダをMapに入れるようにする
        // (再考) フォルダの更新日比較は不要。存在しなければ作成するだけで良い。
        
        // フォルダの存在チェック (API)
        let existingFolderId = null;
        try {
          const folderCheck = Drive.Files.list({
            q: `'${destParentFolderId}' in parents and trashed = false and name = '${file.name.replace(/'/g, "\\'")}' and mimeType = 'application/vnd.google-apps.folder'`,
            fields: 'files(id)',
            supportsAllDrives: true, includeItemsFromAllDrives: true
          });
          if (folderCheck.files && folderCheck.files.length > 0) {
            existingFolderId = folderCheck.files[0].id;
          }
        } catch(e) { /* 無視 */ }

        if (existingFolderId) {
          newDestFolderId = existingFolderId;
        } else {
          // 存在しないので新規作成
          const subFolderResource = { name: file.name, parents: [destParentFolderId], mimeType: 'application/vnd.google-apps.folder' };
          const newSubFolder = Drive.Files.create(subFolderResource, null, { supportsAllDrives: true, fields: 'id' });
          newDestFolderId = newSubFolder.id;
        }
        
        syncRecursively_v3(file.id, newDestFolderId, destFileMap, path);

      } else {
        // --- ファイルの場合 ---
        properties.setProperty('sync_status', `ファイル確認中: ${path}`);
        const destFileInfo = destFileMap.get(path);
        
        let shouldCopy = false;
        if (!destFileInfo) {
          // Case 1: 存在しない
          shouldCopy = true;
        } else {
          // Case 2: 存在する -> 更新日を比較
          const sourceTime = new Date(file.modifiedTime).getTime();
          const destTime = new Date(destFileInfo.modifiedTime).getTime();
          if (sourceTime > destTime) {
            // コピー元が新しい
            try {
              // 古いファイルを削除
              Drive.Files.remove(destFileInfo.id, { supportsAllDrives: true });
            } catch (e) {
              // 削除失敗（例：ゴミ箱にすでにある）してもコピーは続行
            }
            shouldCopy = true;
          }
        }

        if (shouldCopy) {
          properties.setProperty('sync_status', `コピー実行中: ${file.name}`);
          const copyResource = { name: file.name, parents: [destParentFolderId] };
          Drive.Files.copy(copyResource, file.id, { supportsAllDrives: true });
        }
        
        // 進捗を+1
        const currentProcessed = parseInt(properties.getProperty('sync_processedFiles') || '0', 10);
        properties.setProperty('sync_processedFiles', (currentProcessed + 1).toString());
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
}