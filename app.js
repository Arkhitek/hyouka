// リネーム後: app.js (旧: app_jtccm.js)
// 仕様はJTCCM準拠だがファイル名から jtccm 文字列を除去

(function(){
  'use strict';

  // === State ===
  let rawData = [];
  let header = [];
  let envelopeData = null;
  let analysisResults = {};
  let relayoutHandlerAttached = false; // Plotly autoscale対策のイベント重複防止
  // イベントハンドラ参照（重複登録防止用）
  let _plotClickHandler = null;
  let _keydownHandler = null;
  // 包絡線フィット範囲キャッシュ
  let cachedEnvelopeRange = null;

  // === Elements ===
  const gammaInput = document.getElementById('gammaInput');
  const loadInput = document.getElementById('loadInput');
  const wall_length_m = document.getElementById('wall_length_m');
  const test_method = document.getElementById('test_method');
  const alpha_factor = document.getElementById('alpha_factor');
  const envelope_side = document.getElementById('envelope_side');
  const processButton = document.getElementById('processButton');
  const downloadExcelButton = document.getElementById('downloadExcelButton');
  const createShareLinkButton = document.getElementById('createShareLinkButton');
  const clearDataButton = document.getElementById('clearDataButton');
  
  const plotDiv = document.getElementById('plot');
  const pointTooltip = document.getElementById('pointTooltip');
  const undoButton = document.getElementById('undoButton');
  const redoButton = document.getElementById('redoButton');
  const openPointEditButton = null; // ボタンは廃止
  const pointEditDialog = document.getElementById('pointEditDialog');
  const editGammaInput = document.getElementById('edit_gamma');
  const editLoadInput = document.getElementById('edit_load');
  const applyPointEditButton = document.getElementById('applyPointEdit');
  const cancelPointEditButton = document.getElementById('cancelPointEdit');

  // 履歴管理 (Undo/Redo)
  let historyStack = [];
  let redoStack = [];
  const MAX_HISTORY = 100;

  function cloneEnvelope(env){
    return env.map(pt => ({gamma: pt.gamma, Load: pt.Load}));
  }

  function pushHistory(current){
    if(!current) return;
    historyStack.push(cloneEnvelope(current));
    if(historyStack.length > MAX_HISTORY){ historyStack.shift(); }
    redoStack = [];
    updateHistoryButtons();
  }

  function updateHistoryButtons(){
    if(undoButton) undoButton.disabled = historyStack.length <= 1;
    if(redoButton) redoButton.disabled = redoStack.length === 0;
  }

  function performUndo(){
    if(historyStack.length <= 1) return;
    const current = historyStack.pop();
    redoStack.push(current);
    const prev = cloneEnvelope(historyStack[historyStack.length - 1]);
    envelopeData = prev;
  appendLog('Undo: 包絡線を前状態へ戻しました');
  recalculateFromEnvelope(envelopeData);
    window._selectedEnvelopePoint = -1;
    updateHistoryButtons();
  }

  function performRedo(){
    if(redoStack.length === 0) return;
    const next = redoStack.pop();
    historyStack.push(cloneEnvelope(next));
    envelopeData = cloneEnvelope(next);
  appendLog('Redo: 包絡線編集をやり直しました');
  recalculateFromEnvelope(envelopeData);
    window._selectedEnvelopePoint = -1;
    updateHistoryButtons();
  }

  function openPointEditDialog(){
    if(window._selectedEnvelopePoint < 0 || !envelopeData) return;
    const idx = window._selectedEnvelopePoint;
    const pt = envelopeData[idx];
    console.debug('[openPointEditDialog] 開始 idx='+idx+' γ='+pt.gamma+' P='+pt.Load);
    
    // キャンセル用に元値保存（ダイアログ開く前に保存）
    const originalGamma = pt.gamma;
    const originalLoad = pt.Load;
    pointEditDialog.dataset.originalGamma = originalGamma.toString();
    pointEditDialog.dataset.originalLoad = originalLoad.toString();
    
    editGammaInput.value = pt.gamma.toFixed(4);
    editLoadInput.value = pt.Load.toFixed(1);
    
    // ダイアログを選択点と重ならない位置に配置
    pointEditDialog.classList.add('custom-position');
    pointEditDialog.style.display = 'flex';
    const content = document.getElementById('pointEditContent');
    if(content){
      content.style.position = 'absolute';
      // 選択点のスクリーン座標を取得
      const xaxis = plotDiv._fullLayout.xaxis;
      const yaxis = plotDiv._fullLayout.yaxis;
      if(xaxis && yaxis){
        const pointX = xaxis.c2p(pt.gamma); // プロット内でのX座標（ピクセル）
        const pointY = yaxis.c2p(pt.Load);  // プロット内でのY座標（ピクセル）
        const plotRect = plotDiv.getBoundingClientRect();
        
        // 画面上の絶対座標
        const screenX = plotRect.left + pointX;
        const screenY = plotRect.top + pointY;
        
        // ダイアログのサイズ（実測値ベース）
        const dialogWidth = 340;
        const dialogHeight = 220;
        const margin = 60; // 点との間隔を広めに取る
        
        // 点が画面左半分にあれば右側に、右半分にあれば左側に表示
        let left, top;
        const centerX = window.innerWidth / 2;
        const centerY = window.innerHeight / 2;
        
        if(screenX < centerX){
          // 点が左側 → ダイアログを右側に
          left = screenX + margin;
        } else {
          // 点が右側 → ダイアログを左側に
          left = screenX - dialogWidth - margin;
        }
        
        // 点が画面上半分にあれば下側に、下半分にあれば上側に
        if(screenY < centerY){
          // 点が上側 → ダイアログを下側に
          top = screenY + margin;
        } else {
          // 点が下側 → ダイアログを上側に
          top = screenY - dialogHeight - margin;
        }
        
        // 画面外にはみ出さないように調整
        const minMargin = 20;
        left = Math.max(minMargin, Math.min(left, window.innerWidth - dialogWidth - minMargin));
        top = Math.max(minMargin, Math.min(top, window.innerHeight - dialogHeight - minMargin));
        
        content.style.left = left + 'px';
        content.style.top = top + 'px';
        content.style.transform = 'none';
        
        console.debug('[ダイアログ配置] 点位置=('+screenX.toFixed(0)+','+screenY.toFixed(0)+') → ダイアログ位置=('+left+','+top+')');
      } else {
        // フォールバック: 中央表示
        content.style.position = 'absolute';
        content.style.left = '50%';
        content.style.top = '120px';
        content.style.transform = 'translateX(-50%)';
      }
    }
    // 編集中リアルタイム反映
    editGammaInput.oninput = function(){
      const v = parseFloat(editGammaInput.value);
      if(!isNaN(v)){
        envelopeData[idx].gamma = v;
        renderPlot(envelopeData, analysisResults); // 軽量更新: 全描画再生成（簡易）
      }
    };
    editLoadInput.oninput = function(){
      const v = parseFloat(editLoadInput.value);
      if(!isNaN(v)){
        envelopeData[idx].Load = v;
        renderPlot(envelopeData, analysisResults);
      }
    };
    
    // キャンセルボタンのハンドラを上書き（クロージャで現在の値をキャプチャ）
    cancelPointEditButton.onclick = function(){
      if(idx >= 0 && envelopeData && envelopeData[idx]){
        envelopeData[idx].gamma = originalGamma;
        envelopeData[idx].Load = originalLoad;
        console.debug('[キャンセル] 元値に復元: γ='+originalGamma+' P='+originalLoad);
        renderPlot(envelopeData, analysisResults);
      }
      closePointEditDialog();
    };
  }

  function closePointEditDialog(){ 
    pointEditDialog.style.display = 'none'; 
    pointEditDialog.classList.remove('custom-position');
    const content = document.getElementById('pointEditContent');
    if(content){
      content.style.position = '';
      content.style.left = '';
      content.style.top = '';
      content.style.transform = '';
    }
    console.debug('[ダイアログ] 閉じました');
  }

  function applyPointEdit(){
    console.debug('[適用ボタン] クリックされました');
    if(window._selectedEnvelopePoint < 0 || !envelopeData) {
      console.warn('[適用ボタン] 選択点またはデータが無効です');
      return;
    }
    const g = parseFloat(editGammaInput.value);
    const l = parseFloat(editLoadInput.value);
    if(isNaN(g) || isNaN(l)){ 
      alert('数値が不正です'); 
      console.warn('[適用ボタン] 数値が不正です: γ='+g+', P='+l);
      return; 
    }
    envelopeData[window._selectedEnvelopePoint].gamma = g;
    envelopeData[window._selectedEnvelopePoint].Load = l;
    appendLog('点を数値編集しました (γ='+g+', P='+l+')');
    pushHistory(envelopeData);
    console.debug('[適用ボタン] closePointEditDialog()を呼び出します');
    closePointEditDialog();
    // 先に再計算して範囲を再評価
    recalculateFromEnvelope(envelopeData);
    // 描画後、requestAnimationFrameで包絡線範囲を適用して全体フィット化を阻止
    requestAnimationFrame(()=>{
      if(cachedEnvelopeRange){
        Plotly.relayout(plotDiv, {
          'xaxis.autorange': false,
          'yaxis.autorange': false,
          'xaxis.range': cachedEnvelopeRange.xRange,
          'yaxis.range': cachedEnvelopeRange.yRange
        });
      }else{
        fitEnvelopeRanges('点編集後キャッシュ無し');
      }
    });
  }
  
  // キャンセルボタンのハンドラは openPointEditDialog 内で動的に設定されるため、ここでは不要
  
  // 削除ボタン
  const deletePointEditButton = document.getElementById('deletePointEdit');
  if(deletePointEditButton){
    deletePointEditButton.onclick = function(){
      if(window._selectedEnvelopePoint >= 0){
        deleteEnvelopePoint(window._selectedEnvelopePoint, envelopeData);
        window._selectedEnvelopePoint = -1;
        renderPlot(envelopeData, analysisResults);
        closePointEditDialog();
      }
    };
  }
  // 追加ボタン
  const addPointEditButton = document.getElementById('addPointEdit');
  if(addPointEditButton){
    addPointEditButton.onclick = function(){
      if(window._selectedEnvelopePoint >= 0 && envelopeData){
        const idx = window._selectedEnvelopePoint;
        if(idx >= envelopeData.length - 1){
          alert('最後の点の次には追加できません。');
          return;
        }
        // 選択点と次の点の中間値を計算
        const pt1 = envelopeData[idx];
        const pt2 = envelopeData[idx + 1];
        const midGamma = (pt1.gamma + pt2.gamma) / 2;
        const midLoad = (pt1.Load + pt2.Load) / 2;
        
        // 履歴に保存
        pushHistory(envelopeData);
        
        // 新しい点を挿入
        envelopeData.splice(idx + 1, 0, {
          gamma: midGamma,
          Load: midLoad,
          gamma0: midGamma
        });
        
        appendLog('包絡線点を追加しました（γ=' + midGamma.toFixed(6) + ', P=' + midLoad.toFixed(3) + '）');
        renderPlot(envelopeData, analysisResults);
        recalculateFromEnvelope(envelopeData);
        
        // 新しく追加した点を選択
        window._selectedEnvelopePoint = idx + 1;
        
        // ダイアログを閉じて新しい点のダイアログを開く
        closePointEditDialog();
        setTimeout(function(){
          openPointEditDialog();
        }, 100);
      }
    };
  }
  // ダイアログドラッグ移動
  (function enableDialogDrag(){
    const content = document.getElementById('pointEditContent');
    const handle = content ? content.querySelector('.drag-handle') : null;
    if(!content || !handle) return;
    let dragging = false; let startX=0, startY=0, origLeft=0, origTop=0;
    handle.addEventListener('mousedown', function(e){
      dragging = true; startX = e.clientX; startY = e.clientY;
      const rect = content.getBoundingClientRect();
      origLeft = rect.left; origTop = rect.top; content.style.transform='';
      document.body.style.userSelect='none';
    });
    window.addEventListener('mousemove', function(e){
      if(!dragging) return;
      const dx = e.clientX - startX; const dy = e.clientY - startY;
      content.style.left = (origLeft + dx) + 'px';
      content.style.top = (origTop + dy) + 'px';
    });
    window.addEventListener('mouseup', function(){
      dragging = false; document.body.style.userSelect='';
    });
  })();

  // ローカル(file://)でのCORS制約回避用: 組込サンプルCSV（fetch失敗時のフォールバック）
  const BUILTIN_SAMPLE_CSV = `gamma,Load\n
2.28743E-05,0
8.52363E-05,0.42
0.000109903,0.79
0.000129205,1.23
0.000204985,1.46
0.000220261,1.91
0.000272465,2.34
0.000346011,2.77
0.00039815,3.26
0.000458278,3.76
0.000584535,4.43
0.000564155,4.61
0.000711558,5.19
0.000761398,5.77
0.000871873,6.55
0.000961902,7.23
0.00105653,7.69
0.001138636,8.32
0.001247383,8.91
0.001379134,9.76
0.001195178,7.47
0.000964071,5.79
0.000792001,4.24
0.00059316,2.62
0.000385629,1.48
0.000206401,0.45
0.000163458,0
9.06139E-05,-0.14
0.000110605,-0.26
9.68618E-05,-0.39
6.81682E-05,-0.65
5.90237E-05,-0.94
3.57597E-05,-1.27
3.52402E-05,-1.69
-4.86582E-05,-2.17
-0.00011441,-2.73
-0.000192749,-3.36
-0.000284246,-3.79
-0.000332424,-4.16
-0.000347505,-4.27
-0.00041466,-4.73
-0.000511015,-5.35
-0.000574208,-5.81
-0.000660341,-6.39
-0.000715611,-6.79
-0.000777271,-7.31
-0.000868833,-7.67
-0.000896631,-8.13
-0.000980399,-8.69
-0.000882121,-6.85
-0.000740459,-5.25
-0.000600525,-3.77
-0.000398293,-2.34
-0.000252098,-1.47
-0.000246525,-0.73
-0.000139493,-0.17
-0.000166706,0
-0.000100187,0.12
-7.86507E-05,0.22
-7.66633E-05,0.37
-7.2182E-05,0.54
2.81999E-05,0.77
3.56169E-05,1.1
0.000135102,1.64
0.000199893,2.09
0.000252864,2.51
0.000335671,2.94
0.000452537,3.81
0.000494324,3.98
0.000555919,4.56
0.000683644,5.28
0.000828683,6.01
0.00089111,6.8
0.001007014,7.47
0.001136336,8.29
0.00126962,9.2
0.001365014,9.83
0.001170966,7.73
0.00100126,6.24`;

  // === Events ===
  gammaInput.addEventListener('input', handleDirectInput);
  loadInput.addEventListener('input', handleDirectInput);
  processButton.addEventListener('click', processData);
  if(downloadExcelButton) downloadExcelButton.addEventListener('click', downloadExcel);
  if(createShareLinkButton) createShareLinkButton.addEventListener('click', createShareLink);
  clearDataButton.addEventListener('click', clearInputData);
  
  if(undoButton) undoButton.addEventListener('click', performUndo);
  if(redoButton) redoButton.addEventListener('click', performRedo);
  if(openPointEditButton) openPointEditButton.addEventListener('click', openPointEditDialog);
  if(applyPointEditButton) applyPointEditButton.addEventListener('click', applyPointEdit);
  if(cancelPointEditButton) cancelPointEditButton.addEventListener('click', closePointEditDialog);


  function clearInputData(){
    gammaInput.value = '';
    loadInput.value = '';
    rawData = [];
    envelopeData = null;
    analysisResults = {};
    
    processButton.disabled = true;
    if(downloadExcelButton) downloadExcelButton.disabled = true;
    if(createShareLinkButton) createShareLinkButton.disabled = true;
    if(undoButton) undoButton.disabled = true;
  if(redoButton) redoButton.disabled = true;
  if(openPointEditButton) openPointEditButton.disabled = true;
  historyStack = [];
  redoStack = [];
    plotDiv.innerHTML = '';
    // 結果表示リセット
  ['val_pmax','val_py','val_dy','val_K','val_pu','val_dv','val_du','val_mu','val_ds','val_p0_a','val_p0_b','val_p0_c','val_p0_d','val_p0','val_pa','val_magnification'].forEach(id=>{
      const el = document.getElementById(id); if(el) el.textContent='-';
    });
  }

  // 編集モードは廃止

  // === 起動時 sample.csv 自動読込のみ ===
  function loadCsvText(text){
    const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
    const pairs = [];
    const numericRegex = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;
    for(const line of lines){
      const cols = line.split(/,|\t|;/).map(c=>c.trim());
      if(cols.length < 2) continue;
      if(!numericRegex.test(cols[0]) || !numericRegex.test(cols[1])) continue; // header/非数値は除外
      pairs.push([parseFloat(cols[0]), parseFloat(cols[1])]);
    }
    if(pairs.length === 0){
      console.warn('CSVに有効なデータがありません。');
      appendLog('警告: CSVに有効なデータがありません');
      return;
    }
    gammaInput.value = pairs.map(p=>p[0]).join('\n');
    loadInput.value = pairs.map(p=>p[1]).join('\n');
    handleDirectInput();
  }

  function autoLoadSample(){
    fetch('sample.csv', {cache:'no-cache'})
      .then(r => r.ok ? r.text() : Promise.reject(new Error('sample.csvが取得できません')))
      .then(text => loadCsvText(text))
      .catch(err => {
        // file:// でのCORS制約時は組込サンプルへフォールバック
        if(location && location.protocol === 'file:'){
          console.warn('file:// での自動読込を組込サンプルにフォールバックします。詳細:', err.message);
          appendLog('情報: sample.csv 取得失敗 → 組込サンプルを使用 ('+ err.message +')');
          loadCsvText(BUILTIN_SAMPLE_CSV);
        } else {
          console.warn('sample.csv 自動読込失敗:', err.message);
          appendLog('警告: sample.csv 自動読込失敗 ('+ err.message +')');
        }
      });
  }

  autoLoadSample();

  // 包絡線ベースの表示範囲を計算（10%マージン、ゼロ幅回避込み）
  function computeEnvelopeRanges(env){
    if(!env || env.length === 0){
      return { xRange: [-1, 1], yRange: [-1, 1] };
    }
    const xs = env.map(pt => pt.gamma);
    const ys = env.map(pt => pt.Load);
    let minX = Math.min(...xs), maxX = Math.max(...xs);
    let minY = Math.min(...ys), maxY = Math.max(...ys);
    // ゼロ幅のときは小さな幅を与える
    if(minX === maxX){ const pad = Math.max(1e-6, Math.abs(minX)*0.1 || 1e-6); minX -= pad; maxX += pad; }
    if(minY === maxY){ const pad = Math.max(1e-3, Math.abs(minY)*0.1 || 1e-3); minY -= pad; maxY += pad; }
    const mx = (maxX - minX) * 0.1;
    const my = (maxY - minY) * 0.1;
    return { xRange: [minX - mx, maxX + mx], yRange: [minY - my, maxY + my] };
  }

  // 包絡線範囲へフィット（初期描画・Autoscaleボタン・ダブルクリックで共通使用）
  function fitEnvelopeRanges(reason){
    try{
      if(!envelopeData || !envelopeData.length) return;
      const r = computeEnvelopeRanges(envelopeData);
      console.info('[Fit] 包絡線範囲へフィット:', reason || '');
      cachedEnvelopeRange = r; // キャッシュ更新
      Plotly.relayout(plotDiv, {
        'xaxis.autorange': false,
        'yaxis.autorange': false,
        'xaxis.range': r.xRange,
        'yaxis.range': r.yRange
      });
    }catch(err){ console.warn('fitEnvelopeRanges エラー', err); }
  }

  // === Direct Input Handling ===
  function handleDirectInput(){
    const gammaText = gammaInput.value.trim();
    const loadText = loadInput.value.trim();

    if(!gammaText || !loadText) return; // どちらか空なら何もしない

    try {
      // 行単位に分割（カンマ区切り等があっても先に改行を優先）
      const gammaLines = gammaText.split(/\r?\n/);
      const loadLines  = loadText.split(/\r?\n/);

      const pairCount = Math.min(gammaLines.length, loadLines.length);
      const parsed = [];
      let skipped = 0;

      // 厳密な数値判定（単位や文字が付いた行は除外）
      const isNumericString = (s) => /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(s);

      for(let i=0; i<pairCount; i++){
        const gStrRaw = gammaLines[i];
        const lStrRaw = loadLines[i];
        if(gStrRaw == null || lStrRaw == null){
          skipped++; continue;
        }
        const gStr = gStrRaw.trim();
        const lStr = lStrRaw.trim();
        if(!gStr || !lStr){ // 空白行はスキップ
          skipped++; continue;
        }

        // 数値文字列判定：行全体が数値であることを要求
        if(!isNumericString(gStr) || !isNumericString(lStr)){
          skipped++; continue; // 項目名など非数値行は無視
        }

        const gNum = parseFloat(gStr);
        const lNum = parseFloat(lStr);

        parsed.push({
          Load: lNum,
          gamma: gNum,
          gamma0: gNum // 直接入力の場合は補正なし
        });
      }

      if(parsed.length === 0){
        console.warn('有効な数値データがありません。');
        return;
      }

      rawData = parsed;
      header = ['Load', 'gamma', 'gamma0'];
      processButton.disabled = false;

      if(skipped > 0){
        console.info(`非数値または空白行を ${skipped} 行スキップしました。有効データ: ${parsed.length} 行`);
      }

      // 最低3点以上で自動解析（2点以下だと線形近似など不安定）
      if(rawData.length >= 3){
        setTimeout(() => processDataDirect(), 50);
      }
    } catch(err) {
      console.error('データ解析エラー:', err);
      appendLog('データ解析エラー: ' + (err && err.stack ? err.stack : err.message));
    }
  }

  // === Main Processing ===
  function processData(){
    try{
      const L = parseFloat(wall_length_m.value);
      const alpha = parseFloat(alpha_factor.value);
      const side = envelope_side.value;
      const method = test_method.value;

      if(!isFinite(L) || !isFinite(alpha)){
        alert('入力値が不正です。数値を正しく入力してください。');
        return;
      }

      // Direct input - data already has gamma and gamma0
      const dataWithAngles = rawData;

      // Step 2: Generate envelope
      envelopeData = generateEnvelope(dataWithAngles, side);
      if(envelopeData.length === 0){
        alert('包絡線の生成に失敗しました。データを確認してください。');
        return;
      }

      // Step 3: Calculate characteristic points
  analysisResults = calculateJTCCMMetrics(envelopeData, method, L, alpha);

      // Step 4: Render results
      renderPlot(envelopeData, analysisResults);
      renderResults(analysisResults);

  if(downloadExcelButton) downloadExcelButton.disabled = false;
      historyStack = [cloneEnvelope(envelopeData)];
      redoStack = [];
      updateHistoryButtons();
    }catch(error){
      alert('計算エラーが発生しました: ' + error.message);
      console.error(error);
      appendLog('計算エラー: ' + (error && error.stack ? error.stack : error.message));
    }
  }

  // === Direct Input Processing ===
  function processDataDirect(){
    try{
      const L = parseFloat(wall_length_m.value);
      const alpha = parseFloat(alpha_factor.value);
      const side = envelope_side.value;
      const method = test_method.value;

      if(!isFinite(L) || !isFinite(alpha)){
        console.warn('入力値が不正です');
        return;
      }

      // Generate envelope from direct input data
      envelopeData = generateEnvelope(rawData, side);
      if(envelopeData.length === 0){
        console.warn('包絡線の生成に失敗しました');
        return;
      }

      // Calculate characteristic points
  analysisResults = calculateJTCCMMetrics(envelopeData, method, L, alpha);

      // Render results
      renderPlot(envelopeData, analysisResults);
      renderResults(analysisResults);

    if(downloadExcelButton) downloadExcelButton.disabled = false;
    if(createShareLinkButton) createShareLinkButton.disabled = false;
      historyStack = [cloneEnvelope(envelopeData)];
      redoStack = [];
      updateHistoryButtons();
    }catch(error){
      console.error('計算エラー:', error);
      appendLog('計算エラー(自動解析): ' + (error && error.stack ? error.stack : error.message));
    }
  }



  // === Envelope Generation (Section II.3) ===
  function generateEnvelope(data, side){
    // Filter data based on selected side
    let filteredData;
    if(side === 'positive'){
      // Positive side: both gamma and Load must be positive
      filteredData = data.filter(pt => pt.gamma >= 0 && pt.Load >= 0);
    } else {
      // Negative side: both gamma and Load must be negative
      filteredData = data.filter(pt => pt.gamma <= 0 && pt.Load <= 0);
    }
    
    // Do NOT sort - process in original order
    if(filteredData.length === 0) return [];
    
    // Build envelope based on maximum deformation (gamma)
    const env = [];
    let maxAbsGamma = 0;
    
    for(const pt of filteredData){
      const absGamma = Math.abs(pt.gamma);
      
      // Keep point if it has larger deformation than previous maximum
      if(absGamma >= maxAbsGamma){
        maxAbsGamma = absGamma;
        env.push({...pt});
      }
    }
    
    // If envelope is still empty, return the filtered data
    if(env.length === 0 && filteredData.length > 0){
      return filteredData;
    }
    
    // Additional safety check
    if(env.length === 0){
      console.error('包絡線が空です。データを確認してください。', {side, dataCount: data.length, filteredCount: filteredData.length});
    }
    
    return env;
  }

  // === JTCCM Metrics Calculation (Sections III, IV, V) ===
  function calculateJTCCMMetrics(envelope, method, L, alpha){
    const results = {};

    // Determine the sign of the envelope (positive or negative side)
    const envelopeSign = envelope[0] && envelope[0].Load < 0 ? -1 : 1;

    // Find Pmax (Section III.1)
    const Pmax = envelope.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), envelope[0]);
    results.Pmax = Math.abs(Pmax.Load);
    results.Pmax_gamma = Math.abs(Pmax.gamma);

    // Calculate Py using Line Method (Section III.1)
    const Py_result = calculatePy_LineMethod(envelope, results.Pmax);
    results.Py = Py_result.Py;
    results.Py_gamma = Py_result.Py_gamma;
    results.lineI = Py_result.lineI;
    results.lineII = Py_result.lineII;
    results.lineIII = Py_result.lineIII;

    // Calculate Pu and μ using Perfect Elasto-Plastic Model (Section IV)
    const Pu_result = calculatePu_EnergyEquivalent(envelope, results.Py, results.Pmax);
    Object.assign(results, Pu_result);

    // Calculate P0 (Section V.1)
    const P0_result = calculateP0(results, envelope, method);
    Object.assign(results, P0_result);

    // Calculate Pa and Magnification (Section V.2, V.3)
    results.Pa = results.P0 * alpha;
    results.magnification = results.Pa / (L * 1.96);
    results.magnification_rounded = Math.floor(results.magnification * 10) / 10; // Round down to 0.1

    return results;
  }

  // === Py Calculation (Line Method - Section III.1) ===
  function calculatePy_LineMethod(envelope, Pmax){
    const p_max = Pmax;

    // Find points at 0.1, 0.4, 0.9 Pmax (using absolute values)
    const p01 = findPointAtLoad(envelope, 0.1 * p_max);
    const p04 = findPointAtLoad(envelope, 0.4 * p_max);
    const p09 = findPointAtLoad(envelope, 0.9 * p_max);

    if(!p01 || !p04 || !p09) throw new Error('0.1/0.4/0.9 Pmax の点が見つかりません');

    // Use absolute values for gamma as well
    const gamma01 = Math.abs(p01.gamma);
    const gamma04 = Math.abs(p04.gamma);
    const gamma09 = Math.abs(p09.gamma);
    const load01 = Math.abs(p01.Load);
    const load04 = Math.abs(p04.Load);
    const load09 = Math.abs(p09.Load);

    // Line I: 0.1 Pmax - 0.4 Pmax
    const lineI = {
      slope: (load04 - load01) / (gamma04 - gamma01),
      intercept: load01 - ((load04 - load01) / (gamma04 - gamma01)) * gamma01
    };

    // Line II: 0.4 Pmax - 0.9 Pmax
    const lineII = {
      slope: (load09 - load04) / (gamma09 - gamma04),
      intercept: load04 - ((load09 - load04) / (gamma09 - gamma04)) * gamma04
    };

    // Line III: Parallel to Line II, tangent to envelope
    const lineIII = findTangentLine(envelope, lineII.slope);

    // Intersection of Line I and Line III
    const gamma_py = (lineIII.intercept - lineI.intercept) / (lineI.slope - lineIII.slope);
    const Py = lineI.slope * gamma_py + lineI.intercept;

    return { Py, Py_gamma: gamma_py, lineI, lineII, lineIII };
  }

  function findPointAtLoad(envelope, targetLoad){
    for(let i=0; i<envelope.length-1; i++){
      const p1 = envelope[i];
      const p2 = envelope[i+1];
      const abs1 = Math.abs(p1.Load);
      const abs2 = Math.abs(p2.Load);
      
      if(abs1 <= targetLoad && abs2 >= targetLoad){
        const ratio = (targetLoad - abs1) / (abs2 - abs1);
        return {
          Load: p1.Load + (p2.Load - p1.Load) * ratio,
          gamma: p1.gamma + (p2.gamma - p1.gamma) * ratio,
          gamma0: p1.gamma0 + (p2.gamma0 - p1.gamma0) * ratio
        };
      }
    }
    return envelope[envelope.length - 1]; // Fallback
  }

  function findTangentLine(envelope, slope){
    // Find the point where (|Load| - slope * |gamma|) is maximum
    let maxIntercept = -Infinity;
    for(const pt of envelope){
      const intercept = Math.abs(pt.Load) - slope * Math.abs(pt.gamma);
      if(intercept > maxIntercept) maxIntercept = intercept;
    }
    return { slope, intercept: maxIntercept };
  }

  // === Pu and μ Calculation (Energy Equivalent - Section IV) ===
  function calculatePu_EnergyEquivalent(envelope, Py, Pmax){
    // Find δy (gamma where Load = Py on envelope)
    const pt_y = findPointAtLoad(envelope, Py);
    const delta_y = Math.abs(pt_y.gamma);

    // Initial stiffness K
    const K = Py / delta_y;

    // Find δu (Section IV.1 Step 9)
    const delta_u_candidate1 = findDeltaU_08Pmax(envelope, Pmax);
    const delta_u_candidate2 = 1/15; // rad
    const delta_u = Math.min(delta_u_candidate1, delta_u_candidate2);

    // Calculate area S under envelope up to δu
    const S = calculateAreaUnderEnvelope(envelope, delta_u);

    // Solve for Pu using energy equivalence (Section IV.1 Step 11-12)
    // S = Pu * (δu - δv/2), where δv = Pu/K
    // S = Pu * δu - Pu²/(2K)
    // Pu²/(2K) - Pu*δu + S = 0
    // Pu = K*δu - sqrt((K*δu)² - 2*K*S)
    const discriminant = Math.pow(K * delta_u, 2) - 2 * K * S;
    if(discriminant < 0){
      console.warn('Pu計算で判別式が負: discriminant =', discriminant);
      appendLog('警告: Pu計算 判別式<0 のためPyにフォールバック (discriminant='+discriminant.toFixed(6)+')');
      // Fallback: use Py
      return {
        delta_y, K, delta_u, S,
        Pu: Py,
        delta_v: delta_y,
        mu: delta_u / delta_y,
        lineV: {start: {gamma:0, Load:0}, end: {gamma: delta_y, Load: Py}},
        lineVI: {gamma_start: delta_y, gamma_end: delta_u, Load: Py}
      };
    }

    const Pu = K * delta_u - Math.sqrt(discriminant);
    const delta_v = Pu / K;
    const mu = delta_u / delta_v;

    // Lines for visualization
    const lineV = {start: {gamma:0, Load:0}, end: {gamma: delta_v, Load: Pu}};
    const lineVI = {gamma_start: delta_v, gamma_end: delta_u, Load: Pu};

    return { delta_y, K, delta_u, S, Pu, delta_v, mu, lineV, lineVI };
  }

  function findDeltaU_08Pmax(envelope, Pmax){
    const threshold = 0.8 * Pmax;
    let delta_u = Math.abs(envelope[envelope.length - 1].gamma);
    
    // Find first point after Pmax where Load < 0.8Pmax
    let passedMax = false;
    for(const pt of envelope){
      if(Math.abs(pt.Load) >= Pmax * 0.99) passedMax = true;
      if(passedMax && Math.abs(pt.Load) < threshold){
        delta_u = Math.abs(pt.gamma);
        break;
      }
    }
    return delta_u;
  }

  function calculateAreaUnderEnvelope(envelope, delta_limit){
    let area = 0;
    let prev = null;
    for(const pt of envelope){
      const absGamma = Math.abs(pt.gamma);
      if(absGamma > delta_limit){
        if(prev && Math.abs(prev.gamma) < delta_limit){
          // Interpolate to delta_limit
          const ratio = (delta_limit - Math.abs(prev.gamma)) / (absGamma - Math.abs(prev.gamma));
          const load_at_limit = Math.abs(prev.Load) + (Math.abs(pt.Load) - Math.abs(prev.Load)) * ratio;
          area += (delta_limit - Math.abs(prev.gamma)) * (Math.abs(prev.Load) + load_at_limit) / 2;
        }
        break;
      }
      if(prev){
        const dg = absGamma - Math.abs(prev.gamma);
        const avg_load = (Math.abs(prev.Load) + Math.abs(pt.Load)) / 2;
        area += dg * avg_load;
      }
      prev = pt;
    }
    return area;
  }

  // === P0 Calculation (Section V.1) ===
  function calculateP0(results, envelope, method){
    const { Py, Pu, mu, Pmax } = results;

    // (a) Yield strength
    const p0_a = Py;

    // (b) Ductility-based
    let p0_b;
    const denom = 2 * mu - 1;
    if(denom > 0){
      // Pu / (1/ sqrt(2μ - 1)) * 0.2 = 0.2 * Pu * sqrt(2μ - 1)
      p0_b = 0.2 * Pu * Math.sqrt(denom);
    }else{
      // フォールバック（定義域外の場合は Py を採用）
      p0_b = Py;
    }

    // (c) Max strength
    const p0_c = Pmax * (2/3);

    // (d) Specific deformation
    let p0_d;
    if(method === 'loaded'){
      // γ @ 1/120 rad
      const pt = findPointAtGamma(envelope, 1/120, 'gamma');
      p0_d = pt ? Math.abs(pt.Load) : Pmax;
    }else{
      // γ0 @ 1/150 rad
      const pt = findPointAtGamma(envelope, 1/150, 'gamma0');
      p0_d = pt ? Math.abs(pt.Load) : Pmax;
    }

    const P0 = Math.min(p0_a, p0_b, p0_c, p0_d);

    return { p0_a, p0_b, p0_c, p0_d, P0 };
  }

  function findPointAtGamma(envelope, targetGamma, key){
    for(let i=0; i<envelope.length-1; i++){
      const p1 = envelope[i];
      const p2 = envelope[i+1];
      const abs1 = Math.abs(p1[key]);
      const abs2 = Math.abs(p2[key]);
      
      if(abs1 <= targetGamma && abs2 >= targetGamma){
        const ratio = (targetGamma - abs1) / (abs2 - abs1);
        return {
          Load: Math.abs(p1.Load) + (Math.abs(p2.Load) - Math.abs(p1.Load)) * ratio,
          gamma: p1.gamma + (p2.gamma - p1.gamma) * ratio,
          gamma0: p1.gamma0 + (p2.gamma0 - p1.gamma0) * ratio
        };
      }
    }
    return envelope[envelope.length - 1];
  }

  // === Rendering ===
  function renderPlot(envelope, results){
    const { Pmax, Py, Py_gamma, lineI, lineII, lineIII, lineV, lineVI, delta_u, delta_v, p0_a, p0_b, p0_c, p0_d } = results;

  // Draw evaluation overlays on the selected side explicitly
  const envelopeSign = (envelope_side && envelope_side.value === 'negative') ? -1 : 1;

  // Calculate data range for auto-fitting based on envelope (not raw data)
      // 現在の範囲を保持するか、新規計算するか
      let ranges;
      const isDialogOpen = pointEditDialog && pointEditDialog.style.display !== 'none';
      if(isDialogOpen && plotDiv && plotDiv._fullLayout && plotDiv._fullLayout.xaxis && plotDiv._fullLayout.yaxis){
        // ポップアップ表示中は既存の範囲を保持
        ranges = {
          xRange: [plotDiv._fullLayout.xaxis.range[0], plotDiv._fullLayout.xaxis.range[1]],
          yRange: [plotDiv._fullLayout.yaxis.range[0], plotDiv._fullLayout.yaxis.range[1]]
        };
        console.debug('[renderPlot] ポップアップ表示中 - 描画範囲を保持:', ranges);
      } else {
        // 新規計算
        ranges = computeEnvelopeRanges(envelope);
        console.debug('[renderPlot] 描画範囲を新規計算:', ranges);
      }
    
    // 包絡線データを編集可能にするための状態管理
    let editableEnvelope = envelope.map(pt => ({...pt}));    // Original raw data (all points) - showing positive and negative loads
    const trace_rawdata = {
      x: rawData.map(pt => pt.gamma), // rad
      y: rawData.map(pt => pt.Load), // Keep original sign
      mode: 'lines+markers',
      name: '実験データ',
      line: {color: 'lightblue', width: 1},
      marker: {color: 'lightblue', size: 4}
    };

    // Envelope line - keep original sign
    const trace_env = {
      x: editableEnvelope.map(pt => pt.gamma),
      y: editableEnvelope.map(pt => pt.Load), // Keep original sign from filtered data
      mode: 'lines',
      name: '包絡線',
      line: {color: 'blue', width: 2}
    };

    // Envelope points (editable) - click to edit
    const trace_env_points = {
      x: editableEnvelope.map(pt => pt.gamma),
      y: editableEnvelope.map(pt => pt.Load),
      mode: 'markers',
      name: '包絡線点',
      marker: {
        color: editableEnvelope.map((pt, idx) => idx === (window._selectedEnvelopePoint || -1) ? 'red' : 'blue'),
        size: editableEnvelope.map((pt, idx) => idx === (window._selectedEnvelopePoint || -1) ? 14 : 10),
        symbol: 'circle', 
        line: {color: 'white', width: 2}
      },
      hovertemplate: '<b>変形角:</b> %{x:.6f}<br><b>荷重:</b> %{y:.3f}<br><i>クリックで編集、Delキーで削除</i><extra></extra>'
    };

    // Line I, II, III (Py determination)
    const gamma_range = [0, Math.max(...envelope.map(pt => Math.abs(pt.gamma)))]
    ;
    const trace_lineI = makeLine(lineI, gamma_range, 'Line I (0.1-0.4Pmax)', 'orange', envelopeSign);
    const trace_lineIII = makeLine(lineIII, gamma_range, 'Line III (接線)', 'red', envelopeSign);

    // Py point
    const trace_py = {
      x: [Py_gamma * envelopeSign],
      y: [Py * envelopeSign],
      mode: 'markers',
      name: 'Py (降伏耐力)',
      marker: {color: 'green', size: 12, symbol: 'circle'}
    };

    // Perfect elasto-plastic model (Line V, VI)
    const trace_lineV = {
      x: [0, lineV.end.gamma * envelopeSign],
      y: [0, lineV.end.Load * envelopeSign],
      mode: 'lines',
      name: 'Line V (初期剛性)',
      line: {color: 'purple', width: 2, dash: 'dash'}
    };

    const trace_lineVI = {
      x: [lineVI.gamma_start * envelopeSign, lineVI.gamma_end * envelopeSign],
      y: [lineVI.Load * envelopeSign, lineVI.Load * envelopeSign],
      mode: 'lines',
      name: 'Line VI (Pu)',
      line: {color: 'purple', width: 2, dash: 'dash'}
    };

    // Pmax
    const trace_pmax = {
      x: [results.Pmax_gamma * envelopeSign],
      y: [Pmax * envelopeSign],
      mode: 'markers',
      name: 'Pmax',
      marker: {color: 'red', size: 12, symbol: 'star'}
    };

    // P0 criteria lines
  const gamma_max = Math.max(...envelope.map(pt => Math.abs(pt.gamma)));
    const trace_p0_lines = {
      x: [0, gamma_max * envelopeSign, NaN, 0, gamma_max * envelopeSign, NaN, 0, gamma_max * envelopeSign, NaN, 0, gamma_max * envelopeSign],
      y: [p0_a * envelopeSign, p0_a * envelopeSign, NaN, p0_b * envelopeSign, p0_b * envelopeSign, NaN, p0_c * envelopeSign, p0_c * envelopeSign, NaN, p0_d * envelopeSign, p0_d * envelopeSign],
      mode: 'lines',
      name: 'P0基準 (a,b,c,d)',
      line: {color: 'gray', width: 1, dash: 'dot'}
    };

    const layout = {
      title: '荷重-変形関係と評価直線',
      xaxis: {
        title: '変形角 γ (rad)',
        range: ranges.xRange,
        autorange: false
      },
      yaxis: {
        title: '荷重 P (kN)',
        range: ranges.yRange,
        autorange: false
      },
      hovermode: 'closest',
      showlegend: true,
      height: 600,
      annotations: [
        // 終局変位 δu (rad) → Line VI の終点（delta_u の位置）に表示
        {
          x: (lineVI.gamma_end) * envelopeSign,
          y: (lineVI.Load) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `δu=${delta_u.toExponential(2)} rad`,
          showarrow: true,
          ax: 20, ay: -20,
          font: {size: 12, color: 'purple'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'purple', borderwidth: 1
        },
        // 降伏耐力 Py (kN) と 降伏変位 δy (rad) → Py点に表示
        {
          x: (Py_gamma) * envelopeSign,
          y: (Py) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `Py=${Py.toFixed(1)} kN\nδy=${Py_gamma.toExponential(2)} rad`,
          showarrow: true,
          ax: 20, ay: -40,
          font: {size: 12, color: 'green'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'green', borderwidth: 1
        },
        // 終局耐力 Pu (kN) → Line V の終点（delta_v の位置）に表示
        {
          x: (lineV.end.gamma) * envelopeSign,
          y: (lineV.end.Load) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `Pu=${(lineVI.Load).toFixed(1)} kN`,
          showarrow: true,
          ax: 20, ay: -20,
          font: {size: 12, color: 'purple'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'purple', borderwidth: 1
        },
        // 降伏点変位 δv (rad) → Line V の終点（delta_v の位置）に表示
        {
          x: (lineV.end.gamma) * envelopeSign,
          y: (lineV.end.Load) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `δv=${delta_v.toExponential(2)} rad`,
          showarrow: true,
          ax: -30, ay: 20,
          font: {size: 12, color: 'purple'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'purple', borderwidth: 1
        },
        // 最大耐力 Pmax (kN) → Pmax点に表示
        {
          x: (results.Pmax_gamma) * envelopeSign,
          y: (Pmax) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `Pmax=${Pmax.toFixed(1)} kN`,
          showarrow: true,
          ax: 20, ay: -20,
          font: {size: 12, color: 'red'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'red', borderwidth: 1
        }
      ]
    };

  const plotConfig = {
    editable: false,
    displayModeBar: true,
    // デフォルトのAutoscale/Resetを削除（全データへのフィットを防止）
    modeBarButtonsToRemove: ['autoScale2d', 'resetScale2d'],
    // 包絡線範囲へのフィット専用ボタンを追加
    modeBarButtonsToAdd: [
      {
        name: '包絡線にフィット',
        icon: (Plotly && Plotly.Icons && Plotly.Icons.autoscale) ? Plotly.Icons.autoscale : undefined,
        click: function(gd){
          if(envelopeData && envelopeData.length){
            fitEnvelopeRanges('モードバー');
          }
        }
      }
    ]
  };

  Plotly.newPlot(plotDiv, [trace_rawdata, trace_env, trace_env_points, trace_lineI, trace_lineIII, trace_py, trace_lineV, trace_lineVI, trace_pmax, trace_p0_lines], layout, plotConfig)
    .then(function(){
      // 包絡線点の編集機能を実装
      setupEnvelopeEditing(editableEnvelope);
      
      // Autoscale（モードバーやダブルクリック）が発火した場合も包絡線範囲へ調整
      if(!relayoutHandlerAttached){
        plotDiv.on('plotly_relayout', function(e){
          try{
            if(pointEditDialog && pointEditDialog.style.display !== 'none') return;
            if(!e) return;
            // 何らかの理由でautorangeがtrueになった場合、即キャッシュ適用
            if((e['xaxis.autorange'] === true || e['yaxis.autorange'] === true) && cachedEnvelopeRange){
              requestAnimationFrame(()=>{
                Plotly.relayout(plotDiv, {
                  'xaxis.autorange': false,
                  'yaxis.autorange': false,
                  'xaxis.range': cachedEnvelopeRange.xRange,
                  'yaxis.range': cachedEnvelopeRange.yRange
                });
              });
            }
          }catch(err){ console.warn('autoscale再調整エラー', err); }
        });
        // ダブルクリックのリセットでも同様にフィット
        plotDiv.on('plotly_doubleclick', function(){
          try{
            // ポップアップ表示中はダブルクリックリセットをスキップ
            if(pointEditDialog && pointEditDialog.style.display !== 'none'){
              console.debug('[ダブルクリック] ポップアップ表示中のためスキップ');
              return false;
            }
            if(envelopeData && envelopeData.length){
              fitEnvelopeRanges('ダブルクリック');
            }
          }catch(err){ console.warn('doubleclick再調整エラー', err); }
          return false; // 既存のデフォルト動作抑制
        });
        relayoutHandlerAttached = true;
      }
    });
  }

  // === 包絡線点の編集機能 ===
  function setupEnvelopeEditing(editableEnvelope){
    let isDragging = false;
    let dragPointIndex = -1;
    let selectedPointIndex = -1; // Del キー用の選択状態
    // window._selectedEnvelopePoint の初期化をコメントアウト（既存の選択状態を保持）
    if(typeof window._selectedEnvelopePoint === 'undefined'){
      window._selectedEnvelopePoint = -1; // 初回のみ初期化
    }
    
    // 既存クリックハンドラを解除
    if(_plotClickHandler && typeof plotDiv.removeListener === 'function'){
      plotDiv.removeListener('plotly_click', _plotClickHandler);
    }
    // 包絡線点のクリック処理（クリックで即座に数値編集ダイアログを開く）
    _plotClickHandler = function(data){
      console.debug('[plotly_click] event points=', data && data.points ? data.points.length : 0);
      if(!data.points || data.points.length === 0) return;
      const pt = data.points[0];
      console.debug('[plotly_click] curveNumber='+pt.curveNumber+' pointIndex='+pt.pointIndex);
      if(pt.curveNumber === 2){
        selectedPointIndex = pt.pointIndex;
        window._selectedEnvelopePoint = pt.pointIndex;
        // 視覚的に選択反映
        highlightSelectedPoint(editableEnvelope);
        
        // 解析未実行時のフォールバック: envelopeData が存在しなければ自動解析
        if(!envelopeData && rawData && rawData.length >= 3){
          console.info('[plotly_click] 解析前クリック検出 → 自動解析実行');
          processDataDirect(); // 自動解析
          // 解析後に再選択してダイアログ開く（非同期対応）
          setTimeout(function(){
            if(envelopeData && window._selectedEnvelopePoint >= 0){
              openPointEditDialog();
            }
          }, 100);
        } else {
          // ダイアログを開く
          openPointEditDialog();
        }
        console.debug('[plotly_click] ダイアログ表示要求');
        return;
      }
      // 他のトレースをクリック：選択解除
      selectedPointIndex = -1;
      window._selectedEnvelopePoint = -1;
      highlightSelectedPoint(editableEnvelope);
    };
    plotDiv.on('plotly_click', _plotClickHandler);
    
    // ダブルクリックによる点追加やデフォルト操作は許容（別処理はしない）
    
    // ダブルクリックによる新規点追加機能は廃止（仕様変更）
    
    // Delキーで選択中の点を削除
    // 既存のキーリスナーを解除
    if(_keydownHandler){
      document.removeEventListener('keydown', _keydownHandler);
    }
    const handleKeydown = function(e){
      // 編集モード撤廃
      if(e.key === 'Delete' || e.key === 'Del'){
        if(selectedPointIndex >= 0 && selectedPointIndex < editableEnvelope.length){
          deleteEnvelopePoint(selectedPointIndex, editableEnvelope);
          selectedPointIndex = -1;
          window._selectedEnvelopePoint = -1;
        }
        return;
      }
      // Undo/Redo ショートカット
      if((e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z')){
        e.preventDefault();
        performUndo();
        return;
      }
      if((e.ctrlKey || e.metaKey) && (e.key === 'y' || e.key === 'Y')){
        e.preventDefault();
        performRedo();
        return;
      }
      // 'E' で数値編集ダイアログ
      if(!e.ctrlKey && !e.metaKey && (e.key === 'e' || e.key === 'E')){
        e.preventDefault();
        openPointEditDialog();
        return;
      }
    };
    _keydownHandler = handleKeydown;
    document.addEventListener('keydown', _keydownHandler);
    
    // ドラッグ操作は削除（仕様変更）
    
    // マウス移動でドラッグ
    // moveイベントによるドラッグ処理は削除
    
    // マウスアップでドラッグ終了
    // ドラッグ終了/離脱処理も不要
  }
  
  function highlightSelectedPoint(editableEnvelope){
    // 選択された点を赤色で強調表示
    const colors = editableEnvelope.map((pt, idx) => idx === window._selectedEnvelopePoint ? 'red' : 'blue');
    const sizes = editableEnvelope.map((pt, idx) => idx === window._selectedEnvelopePoint ? 14 : 10);
    
    Plotly.restyle(plotDiv, {
      'marker.color': [colors],
      'marker.size': [sizes]
    }, [2]); // trace 2: 包絡線点
    if(openPointEditButton) openPointEditButton.disabled = (window._selectedEnvelopePoint < 0);
  }
  
  function updateEnvelopePlot(editableEnvelope){
    // 包絡線トレース（trace 1）と包絡線点トレース（trace 2）を更新
    Plotly.restyle(plotDiv, {
      x: [editableEnvelope.map(pt => pt.gamma)],
      y: [editableEnvelope.map(pt => pt.Load)]
    }, [1]); // trace 1: 包絡線
    
    const colors = editableEnvelope.map((pt, idx) => idx === window._selectedEnvelopePoint ? 'red' : 'blue');
    const sizes = editableEnvelope.map((pt, idx) => idx === window._selectedEnvelopePoint ? 14 : 10);
    
    Plotly.restyle(plotDiv, {
      x: [editableEnvelope.map(pt => pt.gamma)],
      y: [editableEnvelope.map(pt => pt.Load)],
      'marker.color': [colors],
      'marker.size': [sizes]
    }, [2]); // trace 2: 包絡線点
    // 旧ドラッグ用ツールチップは廃止
    if(pointTooltip){ pointTooltip.style.display = 'none'; }
  }
  
  function deleteEnvelopePoint(pointIndex, editableEnvelope){
    if(editableEnvelope.length <= 2){
      alert('包絡線には最低2点が必要です');
      return;
    }
    // 履歴: 変更前を保存
    pushHistory(editableEnvelope);
    editableEnvelope.splice(pointIndex, 1);
    window._selectedEnvelopePoint = -1; // 選択解除
    updateEnvelopePlot(editableEnvelope);
    recalculateFromEnvelope(editableEnvelope);
    appendLog('包絡線点を削除しました（残り' + editableEnvelope.length + '点）');
    envelopeData = editableEnvelope.map(p=>({...p}));
    updateHistoryButtons();
  }
  
  function addEnvelopePoint(gamma, load, editableEnvelope){
    // 新しい点を適切な位置に挿入（gamma順）
    let insertIdx = editableEnvelope.findIndex(pt => pt.gamma > gamma);
    if(insertIdx < 0) insertIdx = editableEnvelope.length;
    
    editableEnvelope.splice(insertIdx, 0, {
      gamma: gamma,
      Load: load,
      gamma0: gamma // 簡易的に同値
    });
    
    updateEnvelopePlot(editableEnvelope);
    recalculateFromEnvelope(editableEnvelope);
    appendLog('包絡線点を追加しました（γ=' + gamma.toFixed(6) + ', P=' + load.toFixed(3) + '）');
  }
  
  function addEnvelopePointAtNearestSegment(clickX, clickY, xData, yData, editableEnvelope, xaxis, yaxis){
    // クリック位置から最も近い包絡線セグメント（2点間）を見つける
    let minDist = Infinity;
    let nearestSegmentIdx = 0;
    
    for(let i = 0; i < editableEnvelope.length - 1; i++){
      const p1 = editableEnvelope[i];
      const p2 = editableEnvelope[i + 1];
      
      const x1 = xaxis.c2p(p1.gamma);
      const y1 = yaxis.c2p(p1.Load);
      const x2 = xaxis.c2p(p2.gamma);
      const y2 = yaxis.c2p(p2.Load);
      
      // 線分への最短距離を計算
      const dist = pointToSegmentDistance(clickX, clickY, x1, y1, x2, y2);
      
      if(dist < minDist){
        minDist = dist;
        nearestSegmentIdx = i;
      }
    }
    
    // 最寄りセグメントの中点に新しい点を追加
    const p1 = editableEnvelope[nearestSegmentIdx];
    const p2 = editableEnvelope[nearestSegmentIdx + 1];
    const midGamma = (p1.gamma + p2.gamma) / 2;
    const midLoad = (p1.Load + p2.Load) / 2;
    
    // 履歴: 変更前を保存
    pushHistory(editableEnvelope);
    editableEnvelope.splice(nearestSegmentIdx + 1, 0, {
      gamma: midGamma,
      Load: midLoad,
      gamma0: midGamma
    });
    
    updateEnvelopePlot(editableEnvelope);
    recalculateFromEnvelope(editableEnvelope);
    appendLog('包絡線点を追加しました（γ=' + midGamma.toFixed(6) + ', P=' + midLoad.toFixed(3) + '）');
    envelopeData = editableEnvelope.map(p=>({...p}));
    updateHistoryButtons();
  }
  
  function pointToSegmentDistance(px, py, x1, y1, x2, y2){
    // 点(px, py)から線分(x1,y1)-(x2,y2)への最短距離
    const dx = x2 - x1;
    const dy = y2 - y1;
    const lengthSq = dx * dx + dy * dy;
    
    if(lengthSq === 0){
      // 線分が点の場合
      return Math.sqrt((px - x1) * (px - x1) + (py - y1) * (py - y1));
    }
    
    // 線分上の最近点のパラメータt (0 <= t <= 1)
    let t = ((px - x1) * dx + (py - y1) * dy) / lengthSq;
    t = Math.max(0, Math.min(1, t));
    
    const closestX = x1 + t * dx;
    const closestY = y1 + t * dy;
    
    return Math.sqrt((px - closestX) * (px - closestX) + (py - closestY) * (py - closestY));
  }
  
  function recalculateFromEnvelope(editableEnvelope){
    try{
      // 編集後の包絡線から特性値を再計算
      envelopeData = editableEnvelope.map(pt => ({...pt}));
      
      const L = parseFloat(wall_length_m.value);
      const alpha = parseFloat(alpha_factor.value);
      const method = test_method.value;
      
      if(!isFinite(L) || !isFinite(alpha)) return;
      
      analysisResults = calculateJTCCMMetrics(envelopeData, method, L, alpha);
      renderResults(analysisResults);
      
      // 評価直線などを再描画
      renderPlot(envelopeData, analysisResults);
      
      appendLog('包絡線編集に基づき特性値を再計算しました');
    }catch(err){
      console.error('再計算エラー:', err);
      appendLog('再計算エラー: ' + (err && err.message ? err.message : err));
    }
  }

  function makeLine(lineObj, gamma_range, name, color, sign = 1){
    const x = gamma_range.map(g => g * sign);
    const y = gamma_range.map(g => (lineObj.slope * g + lineObj.intercept) * sign);
    return {
      x, y,
      mode: 'lines',
      name,
      line: {color, width: 1, dash: 'dash'}
    };
  }

  function renderResults(r){
    document.getElementById('val_pmax').textContent = r.Pmax.toFixed(3);
    document.getElementById('val_py').textContent = r.Py.toFixed(3);
    document.getElementById('val_dy').textContent = (r.delta_y).toFixed(5) + ' rad';
    document.getElementById('val_K').textContent = r.K.toFixed(2);
    document.getElementById('val_pu').textContent = r.Pu.toFixed(3);
    document.getElementById('val_dv').textContent = (r.delta_v).toFixed(5) + ' rad';
    document.getElementById('val_du').textContent = (r.delta_u).toFixed(5) + ' rad';
    document.getElementById('val_mu').textContent = r.mu.toFixed(3);
    // 構造特性係数 Ds = 1 / sqrt(2μ - 1)
    let Ds = '-';
    if(r.mu && r.mu > 0.5){ // 2μ-1 > 0 の領域のみ算出（μ>0.5）
      const denom = Math.sqrt(2 * r.mu - 1);
      if(denom > 0){ Ds = (1 / denom).toFixed(3); }
    }
    const dsEl = document.getElementById('val_ds');
    if(dsEl) dsEl.textContent = Ds;

    document.getElementById('val_p0_a').textContent = r.p0_a.toFixed(3);
    document.getElementById('val_p0_b').textContent = r.p0_b.toFixed(3);
    document.getElementById('val_p0_c').textContent = r.p0_c.toFixed(3);
    document.getElementById('val_p0_d').textContent = r.p0_d.toFixed(3);
    document.getElementById('val_p0').textContent = r.P0.toFixed(3);

    document.getElementById('val_pa').textContent = r.Pa.toFixed(3);
    document.getElementById('val_magnification').textContent = r.magnification_rounded.toFixed(1) + ' 倍';
  }

  async function downloadExcel(){
    if(!window.ExcelJS){
      alert('ExcelJSライブラリが読み込まれていません');
      return;
    }
    try{
      let wb = null;
      // ネイティブチャート対応: 明示的に有効化された場合のみ template.xlsx を読み込み
      if(window.APP_CONFIG && window.APP_CONFIG.useExcelTemplate){
        try{
          const resp = await fetch('template.xlsx', {cache:'no-cache'});
          if(resp.ok){
            const buf = await resp.arrayBuffer();
            wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buf);
            appendLog('情報: template.xlsx を使用してExcelを生成');
          }else{
            appendLog('情報: template.xlsx が見つかりません (resp='+resp.status+')');
          }
        }catch(e){
          appendLog('情報: template.xlsx 読込不可 (' + (e && e.message ? e.message : e) + ')');
        }
      }
      if(!wb){
        wb = new ExcelJS.Workbook();
        wb.creator = 'hyouka-app';
        wb.created = new Date();
      }

      // 1) 解析結果シート
      let wsSummary = wb.getWorksheet('Summary');
      if(!wsSummary) wsSummary = wb.addWorksheet('Summary');
      const r = analysisResults;
      wsSummary.addRow(['項目','値','単位']);
      const rows = [
        ['最大耐力 Pmax', r.Pmax, 'kN'],
        ['降伏耐力 Py', r.Py, 'kN'],
        ['降伏変位 δy', r.delta_y, 'rad'],
        ['初期剛性 K', r.K, 'kN/rad'],
        ['終局耐力 Pu', r.Pu, 'kN'],
        ['終局変位 δu', r.delta_u, 'rad'],
        ['塑性率 μ', r.mu, ''],
        ['P0(a) 降伏耐力', r.p0_a, 'kN'],
        ['P0(b) 靭性基準', r.p0_b, 'kN'],
        ['P0(c) 最大耐力基準', r.p0_c, 'kN'],
        ['P0(d) 特定変形時', r.p0_d, 'kN'],
        ['短期基準せん断耐力 P0', r.P0, 'kN'],
        ['短期許容せん断耐力 Pa', r.Pa, 'kN'],
        ['壁倍率', r.magnification_rounded, '倍']
      ];
      rows.forEach(row => wsSummary.addRow(row));
      wsSummary.columns.forEach(col => { col.width = 22; });
      // 数値フォーマット適用
      for(let i=2;i<=wsSummary.rowCount;i++){
        const label = wsSummary.getCell(i,1).value;
        const cell = wsSummary.getCell(i,2);
        if(typeof cell.value !== 'number') continue;
        if(/rad/.test(wsSummary.getCell(i,3).value)) cell.numFmt = '0.000000';
        else if(label === '初期剛性 K') cell.numFmt = '#,##0.00';
        else if(label === '壁倍率') cell.numFmt = '0.0';
        else cell.numFmt = '#,##0.000';
      }

      // 2) 入力データシート
      let wsInput = wb.getWorksheet('InputData');
      if(!wsInput) wsInput = wb.addWorksheet('InputData');
      // ヘッダ再設定と既存データクリア
      wsInput.spliceRows(1, wsInput.rowCount, ['gamma','Load']);
      rawData.forEach(pt => wsInput.addRow([pt.gamma, pt.Load]));
      wsInput.columns.forEach(c=> c.width = 18);
      for(let i=2;i<=wsInput.rowCount;i++){
        const cg = wsInput.getCell(i,1); if(typeof cg.value==='number') cg.numFmt='0.000000';
        const cp = wsInput.getCell(i,2); if(typeof cp.value==='number') cp.numFmt='0.000';
      }

      // 3) 包絡線シート
      let wsEnv = wb.getWorksheet('Envelope');
      if(!wsEnv) wsEnv = wb.addWorksheet('Envelope');
      wsEnv.spliceRows(1, wsEnv.rowCount, ['gamma','Load']);
      (envelopeData||[]).forEach(pt => wsEnv.addRow([pt.gamma, pt.Load]));
      wsEnv.columns.forEach(c=> c.width = 18);
      for(let i=2;i<=wsEnv.rowCount;i++){
        const cg = wsEnv.getCell(i,1); if(typeof cg.value==='number') cg.numFmt='0.000000';
        const cp = wsEnv.getCell(i,2); if(typeof cp.value==='number') cp.numFmt='0.000';
      }

      // 4) グラフシート (画像埋込み)
      // Chartシート: テンプレートがあれば既存を活用。無ければ画像埋め込みの代替。
      let wsChart = wb.getWorksheet('Chart');
      if(!wsChart){
        wsChart = wb.addWorksheet('Chart');
        wsChart.getCell('A1').value = '荷重-変形関係グラフ';
        wsChart.getRow(1).font = {bold:true};
        const pngDataUrl = await Plotly.toImage(plotDiv, {format:'png', width:1200, height:700});
        const base64 = pngDataUrl.replace(/^data:image\/png;base64,/, '');
        const imageId = wb.addImage({base64, extension:'png'});
        wsChart.addImage(imageId, { tl: {col:0, row:2}, ext: {width: 900, height: 520} });
      }

      // 仕上げ: 自動フィルタやスタイル軽微調整
      wsSummary.getRow(1).font = {bold:true};
      wsInput.getRow(1).font = {bold:true};
      wsEnv.getRow(1).font = {bold:true};

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'Results.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }catch(err){
      console.error('Excel出力エラー:', err);
      alert('Excelの生成に失敗しました。');
      appendLog('Excel出力エラー: ' + (err && err.stack ? err.stack : err.message));
    }
  }

  function appendLog(message){
    // ログ機能は無効化（コンソールのみに出力）
    console.log('[LOG]', message);
  }

    // === Share Link Functions ===
    function createShareLink(){
      if(!analysisResults || !envelopeData){
        alert('解析結果がありません。まずデータを入力して解析を実行してください。');
        return;
      }

      try{
        // Serialize input data and parameters
        if(!wall_length_m || !test_method || !alpha_factor || !envelope_side){
          throw new Error('必須入力要素が取得できません (IDの変更やDOM未構築の可能性)');
        }
        const shareData = {
          data: rawData.map(d => [d.gamma, d.Load]), // [[gamma, Load], ...]
          wall_length: parseFloat(wall_length_m.value) || 1.0,
          test_method: (test_method.value || '').trim(),
          alpha: parseFloat(alpha_factor.value) || 1.0,
          side: (envelope_side.value || '').trim()
        };

        // Encode to base64 JSON
        const jsonStr = JSON.stringify(shareData);
        const base64Data = btoa(unescape(encodeURIComponent(jsonStr)));

        // Create URL with query parameter
        const url = new URL(window.location.href.split('?')[0]);
        url.searchParams.set('share', base64Data);

        // Copy to clipboard
        navigator.clipboard.writeText(url.toString()).then(() => {
          alert('共有リンクをクリップボードにコピーしました。\n\nリンクを共有すると、同じ解析結果を表示できます。');
        }).catch(err => {
          // Fallback: show URL in prompt
          prompt('共有リンクをコピーしてください:', url.toString());
        });

        appendLog('共有リンク作成: ' + url.toString());
      }catch(error){
        console.error('共有リンク作成エラー:', error);
        alert('共有リンクの作成に失敗しました。');
      }
    }

    function loadFromSharedLink(){
      try{
        const urlParams = new URLSearchParams(window.location.search);
        const shareParam = urlParams.get('share');

        if(!shareParam){
          return; // No shared link
        }

        // Decode base64 JSON
        const jsonStr = decodeURIComponent(escape(atob(shareParam)));
        const shareData = JSON.parse(jsonStr);

        // Validate data structure
        if(!shareData.data || !Array.isArray(shareData.data)){
          throw new Error('Invalid share data format');
        }

        // Populate input fields
  if(wall_length_m) wall_length_m.value = shareData.wall_length || 1.0;
  if(test_method) test_method.value = shareData.test_method || 'monotonic';
  if(alpha_factor) alpha_factor.value = shareData.alpha || 1.0;
  if(envelope_side) envelope_side.value = shareData.side || 'positive';

        // Populate data table
        rawData = shareData.data.map(([gamma, Load]) => ({ gamma, Load }));
        gammaInput.value = rawData.map(d => d.gamma).join('\n');
        loadInput.value = rawData.map(d => d.Load).join('\n');

        // Auto-process
        processData();

        appendLog('共有リンクからデータを読み込みました');
        alert('共有リンクからデータを読み込みました。');
      }catch(error){
        console.error('共有リンク読み込みエラー:', error);
        alert('共有リンクの読み込みに失敗しました。URLが正しいか確認してください。');
      }
    }

    // Load from shared link on page load
    window.addEventListener('DOMContentLoaded', () => {
      loadFromSharedLink();
    });
})();
