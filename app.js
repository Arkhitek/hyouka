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
  const specific_deformation = document.getElementById('specific_deformation');
  const alpha_factor = document.getElementById('alpha_factor');
  const max_ultimate_deformation = document.getElementById('max_ultimate_deformation');
  const c0_factor = document.getElementById('c0_factor');
  const wall_preset = document.getElementById('wall_preset');
  const envelope_side = document.getElementById('envelope_side');
  const specimen_name = document.getElementById('specimen_name');
  // 手動解析ボタンは廃止
  const processButton = null;
  const downloadExcelButton = document.getElementById('downloadExcelButton');
  const generatePdfButton = document.getElementById('generatePdfButton');
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

  // === Utilities ===
  // 角度[rad]を 1/N 表記の文字列へ変換（Nは四捨五入した整数）
  function formatReciprocal(rad){
    const v = Number(rad);
    if(!isFinite(v) || v <= 0) return '-';
    const denom = Math.round(1 / v);
    if(!isFinite(denom) || denom <= 0) return '-';
    return '1/' + denom.toLocaleString('ja-JP');
  }

  // === Preset application ===
  function applyWallPreset(code){
    if(!code) return;
    // 定義: specific (1/N), ultimate (1/N), c0
    const map = {
      wood_loaded:   { specific:120, ultimate:15, c0:0.2 },
      wood_tierod:   { specific:150, ultimate:15, c0:0.2 },
      lgs_true:      { specific:200, ultimate:30, c0:0.3 },
      lgs_apparent:  { specific:120, ultimate:30, c0:0.3 }
    };
    const preset = map[code];
    if(!preset) return;
    specific_deformation.value = preset.specific;
    max_ultimate_deformation.value = preset.ultimate;
    c0_factor.value = preset.c0;
    appendLog('対象耐力壁プリセット適用: '+code+' → 特定変形1/'+preset.specific+', 最大終局1/'+preset.ultimate+', C0='+preset.c0);
  }

  if(wall_preset){
    wall_preset.addEventListener('change', e => {
      applyWallPreset(e.target.value);
      // プリセット変更時も自動解析
      if(rawData && rawData.length >= 3){ scheduleAutoRun(); }
    });
  }

  // 自動解析スケジューラ（タイプ中の過剰実行を防止）
  let _autoRunTimer = null;
  function scheduleAutoRun(delay=150){
    if(_autoRunTimer) clearTimeout(_autoRunTimer);
    _autoRunTimer = setTimeout(() => {
      try{ processDataDirect(); }catch(e){ console.warn('auto-run error', e); }
    }, delay);
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

    const originalGamma = pt.gamma;
    const originalLoad = pt.Load;
    pointEditDialog.dataset.originalGamma = originalGamma.toString();
    pointEditDialog.dataset.originalLoad = originalLoad.toString();

    editGammaInput.value = pt.gamma.toFixed(4);
    editLoadInput.value = pt.Load.toFixed(1);

    // 右端固定表示（CSS custom-positionが制御）
    pointEditDialog.classList.add('custom-position');
    pointEditDialog.style.display = 'flex';

    // 編集中リアルタイム反映
    editGammaInput.oninput = function(){
      const v = parseFloat(editGammaInput.value);
      if(!isNaN(v)){
        envelopeData[idx].gamma = v;
        renderPlot(envelopeData, analysisResults);
      }
    };
    editLoadInput.oninput = function(){
      const v = parseFloat(editLoadInput.value);
      if(!isNaN(v)){
        envelopeData[idx].Load = v;
        renderPlot(envelopeData, analysisResults);
      }
    };

    cancelPointEditButton.onclick = function(){
      if(idx >= 0 && envelopeData && envelopeData[idx]){
        envelopeData[idx].gamma = originalGamma;
        envelopeData[idx].Load = originalLoad;
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
      content.style.margin = '';
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
        safeRelayout(plotDiv, {
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
      origLeft = rect.left; origTop = rect.top;
      // ドラッグ開始時にabsolute配置に切り替え
      content.style.position = 'absolute';
      content.style.left = origLeft + 'px';
      content.style.top = origTop + 'px';
      content.style.margin = '0';
      content.style.transform = '';
      document.body.style.userSelect='none';
      e.preventDefault(); // デフォルト動作を抑制
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
  // 手動ボタン削除済み: processButton クリックイベント不要

  // パラメータ変更時の自動解析
  const autoInputs = [wall_length_m, specific_deformation, alpha_factor, max_ultimate_deformation, c0_factor];
  autoInputs.forEach(el => {
    if(!el) return;
    el.addEventListener('input', () => { if(rawData && rawData.length>=3) scheduleAutoRun(); });
    el.addEventListener('change', () => { if(rawData && rawData.length>=3) scheduleAutoRun(); });
  });
  if(envelope_side){
    envelope_side.addEventListener('change', () => { if(rawData && rawData.length>=3) scheduleAutoRun(); });
  }
  if(downloadExcelButton) downloadExcelButton.addEventListener('click', downloadExcel);
  if(generatePdfButton) generatePdfButton.addEventListener('click', generatePdfReport);
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
    
  // 自動解析化に伴い、旧ボタンの状態管理は不要
  if(downloadExcelButton) downloadExcelButton.disabled = true;
  if(generatePdfButton) generatePdfButton.disabled = true;
    if(undoButton) undoButton.disabled = true;
  if(redoButton) redoButton.disabled = true;
  if(openPointEditButton) openPointEditButton.disabled = true;
  historyStack = [];
  redoStack = [];
    plotDiv.innerHTML = '';
    // 結果表示リセット
  ['val_pmax','val_py','val_dy','val_K','val_pu','val_dv','val_du','val_mu','val_ds','val_p0_a','val_p0_b','val_p0_c','val_p0_d','val_p0','val_pa','val_pa_per_m','val_pu_per_m','val_magnification'].forEach(id=>{
      const el = document.getElementById(id); if(el) el.textContent='-';
    });
  }

  // === PDF Generation ===
  async function generatePdfReport(){
    try{
      if(!analysisResults || !envelopeData || !envelopeData.length){
        alert('解析結果がありません');
        return;
      }
      const specimen = (specimen_name && specimen_name.value ? specimen_name.value.trim() : 'testname');
      // Ensure jsPDF and html2canvas
      const { jsPDF } = window.jspdf || {};
      if(!jsPDF){
        alert('jsPDFライブラリが読み込まれていません');
        return;
      }
      if(typeof html2canvas === 'undefined'){
        alert('html2canvasライブラリが読み込まれていません');
        return;
      }

      // Create temporary container for PDF content
      const container = document.createElement('div');
      container.style.cssText = 'position:absolute; left:-9999px; top:0; width:800px; background:white; padding:20px; font-family:sans-serif;';
      document.body.appendChild(container);

      // Build HTML content
      const r = analysisResults;
      const fmt1 = (v) => formatReciprocal(v);
      container.innerHTML = `
        <div style="text-align:center; margin-bottom:15px;">
          <h1 style="font-size:24px; margin:10px 0;">耐力壁性能評価レポート</h1>
          <p style="font-size:12px; color:#666; margin:5px 0;">生成日時: ${new Date().toISOString().replace('T',' ').substring(0,19)}</p>
          <p style="font-size:12px; color:#333; margin:5px 0;">試験体名称: ${specimen.replace(/</g,'&lt;')}</p>
        </div>
        <div id="pdf-plot" style="width:100%; height:400px; margin-bottom:20px;"></div>
        <div style="display:flex; gap:20px;">
          <div style="flex:1;">
            <h3 style="font-size:16px; margin:10px 0; border-bottom:2px solid #333; padding-bottom:5px;">入力パラメータ</h3>
            <table style="width:100%; font-size:12px; border-collapse:collapse; table-layout:fixed;">
              <colgroup>
                <col style="width:60%">
                <col style="width:40%">
              </colgroup>
              <tr style="border-bottom:1px solid #ccc;"><td style="padding:6px 8px;">壁長さ L (m)</td><td style="text-align:right; padding:6px 8px;">${wall_length_m.value}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">特定変形角</td><td style="text-align:right; padding:6px 8px;">1/${Number(specific_deformation.value).toLocaleString('ja-JP')}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">最大終局変位</td><td style="text-align:right; padding:6px 8px;">1/${Number(max_ultimate_deformation.value).toLocaleString('ja-JP')}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">C0</td><td style="text-align:right; padding:6px 8px;">${c0_factor.value}</td></tr>
              <tr style="border-bottom:1px solid #ccc;"><td style="padding:6px 8px;">α</td><td style="text-align:right; padding:6px 8px;">${alpha_factor.value}</td></tr>
            </table>
          </div>
          <div style="flex:1;">
            <h3 style="font-size:16px; margin:10px 0; border-bottom:2px solid #333; padding-bottom:5px;">計算結果</h3>
            <table style="width:100%; font-size:12px; border-collapse:collapse; table-layout:fixed;">
              <colgroup>
                <col style="width:60%">
                <col style="width:40%">
              </colgroup>
              <tr style="border-bottom:1px solid #ccc;"><td style="padding:6px 8px;">Pmax (kN)</td><td style="text-align:right; padding:6px 8px;">${r.Pmax?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">Py (kN)</td><td style="text-align:right; padding:6px 8px;">${r.Py?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">Pu (kN)</td><td style="text-align:right; padding:6px 8px;">${r.Pu?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">Pu (kN/m)</td><td style="text-align:right; padding:6px 8px;">${(function(){const L=parseFloat(wall_length_m.value);return (isFinite(L)&&L>0&&r.Pu)?(r.Pu/L).toFixed(3):'-';})()}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">δv</td><td style="text-align:right; padding:6px 8px;">${fmt1(r.delta_v)}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">δu</td><td style="text-align:right; padding:6px 8px;">${fmt1(r.delta_u)}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">μ</td><td style="text-align:right; padding:6px 8px;">${r.mu?.toFixed(2) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">Ds</td><td style="text-align:right; padding:6px 8px;">${r.mu && r.mu>0 ? (1/Math.sqrt(2*r.mu-1)).toFixed(3) : '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">P0(a)</td><td style="text-align:right; padding:6px 8px;">${r.p0_a?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">P0(b)</td><td style="text-align:right; padding:6px 8px;">${r.p0_b?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">P0(c)</td><td style="text-align:right; padding:6px 8px;">${r.p0_c?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">P0(d)</td><td style="text-align:right; padding:6px 8px;">${r.p0_d?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">P0</td><td style="text-align:right; padding:6px 8px;">${r.P0?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">Pa (kN)</td><td style="text-align:right; padding:6px 8px;">${r.Pa?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #ccc;"><td style="padding:6px 8px;">Pa (kN/m)</td><td style="text-align:right; padding:6px 8px;">${(function(){const L=parseFloat(wall_length_m.value);return (isFinite(L)&&L>0&&r.Pa)?(r.Pa/L).toFixed(3):'-';})()}</td></tr>
              <tr style="border-bottom:1px solid #ccc;"><td style="padding:6px 8px; font-weight:bold;">壁倍率</td><td style="text-align:right; padding:6px 8px; font-weight:bold;">${r.magnification_rounded?.toFixed(1) ?? '-'}</td></tr>
            </table>
          </div>
        </div>
        <div style="text-align:center; font-size:10px; color:#666; margin-top:20px;">© Arkhitek / Generated by 耐力壁性能評価プログラム</div>
      `;

      // Render Plotly graph to temporary container
      const pdfPlotDiv = container.querySelector('#pdf-plot');
      await Plotly.newPlot(pdfPlotDiv, plotDiv.data, {
        ...plotDiv.layout,
        width: 760,
        height: 400,
        margin: {l:60, r:20, t:40, b:60}
      }, {displayModeBar: false});

      // Convert to image using html2canvas
      const canvas = await html2canvas(container, {
        scale: 2,
        backgroundColor: '#ffffff',
        logging: false
      });

      // Remove temporary container
      document.body.removeChild(container);

      // Create PDF
      const doc = new jsPDF({orientation:'portrait', unit:'mm', format:'a4'});
      const imgData = canvas.toDataURL('image/png');
      const pageW = 210;
      const pageH = 297;
      const imgW = pageW;
      const imgH = (canvas.height * pageW) / canvas.width;
      
      // Scale to fit page if necessary
      if(imgH > pageH - 20){
        const scale = (pageH - 20) / imgH;
        doc.addImage(imgData, 'PNG', 0, 10, imgW * scale, imgH * scale);
      } else {
        doc.addImage(imgData, 'PNG', 0, 10, imgW, imgH);
      }

      const pdfFileName = `Report_${specimen.replace(/[^a-zA-Z0-9_\-一-龥ぁ-んァ-ヶ]/g,'_')}.pdf`;
      doc.save(pdfFileName);
      appendLog('PDFレポートを生成しました');
    }catch(err){
      console.error('PDF生成エラー:', err);
      alert('PDF生成に失敗しました: ' + (err && err.message ? err.message : err));
      appendLog('PDF生成エラー: ' + (err && err.stack ? err.stack : err));
    }
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

  // Plotly.relayout の安全ラッパ：引数検証と失敗時の抑止・診断
  function safeRelayout(gd, updates){
    try{
      if(typeof updates === 'string'){
        // 第3引数が必要なケースは呼び出し側の責務。
        // 誤用を避けるため、この分岐ではそのまま通さず警告して無視する。
        console.warn('[safeRelayout] 文字列キーのみの呼び出しはサポートしていません:', updates);
        return Promise.resolve();
      }
      if(!updates || typeof updates !== 'object' || Array.isArray(updates)){
        console.warn('[safeRelayout] invalid updates. expect plain object. got =', updates);
        return Promise.resolve();
      }
      const p = Plotly.relayout(gd, updates);
      // Promise 対応: 失敗を握りつぶして未処理拒否を防ぐ
      if(p && typeof p.then === 'function'){
        return p.catch(err => {
          console.warn('[safeRelayout] relayout rejected:', err);
        });
      }
      return Promise.resolve();
    }catch(err){
      console.warn('[safeRelayout] wrapper error', err);
      return Promise.resolve();
    }
  }

  // 軽量なグローバル抑止: 理由が undefined の未処理Promise拒否のみ握りつぶす（本番でも常時有効）
  // Plotly 内部の一部経路で reject(undefined) が発生するため、実用上のノイズを抑える。
  if(typeof window !== 'undefined' && window.addEventListener){
    window.addEventListener('unhandledrejection', function(ev){
      try{
        if(!ev) return;
        if(typeof ev.reason === 'undefined'){
          // デバッグが必要な場合は ?debug=layout を使用（詳細パッチが有効化される）
          ev.preventDefault();
        }
      }catch(_){/* noop */}
    });
  }

  // デバッグフラグ (?debug=layout または APP_CONFIG.debugLayout) が true のときのみ詳細パッチを適用
  (function patchGlobalRelayout(){
    try{
      if(!window.Plotly) return;
      if(window.__PLOTLY_RELAYOUT_PATCHED__) return;
      function isLayoutDebug(){
        try{ const u=new URL(window.location.href); if(u.searchParams.get('debug')==='layout') return true; }catch(_){/*noop*/}
        try{ if(window.APP_CONFIG && window.APP_CONFIG.debugLayout===true) return true; }catch(_){/*noop*/}
        return false;
      }
      const DEBUG = isLayoutDebug();
      if(!DEBUG){ return; } // 本番は何もしない

      const origRelayout = window.Plotly.relayout;
      if(typeof origRelayout !== 'function') return;
      window.Plotly.relayout = function patchedRelayout(gd, a, b){
        try{
          // 文字列キー指定のときは値が未指定なら無視
          if(typeof a === 'string'){
            if(typeof b === 'undefined'){
              console.warn('[patch.relayout] string key without value. ignore:', a);
              return Promise.resolve();
            }
            const pr = origRelayout.call(window.Plotly, gd, a, b);
            return (pr && typeof pr.then === 'function') ? pr.catch(err => {
              console.warn('[patch.relayout] rejected (string key):', a, b, err);
            }) : Promise.resolve();
          }
          // オブジェクト以外は無視（undefined reject を回避）
          if(!a || typeof a !== 'object' || Array.isArray(a)){
            console.warn('[patch.relayout] invalid updates (expect plain object). got =', a);
            return Promise.resolve();
          }
          const pr = origRelayout.call(window.Plotly, gd, a);
          return (pr && typeof pr.then === 'function') ? pr.catch(err => {
            console.warn('[patch.relayout] rejected (object updates):', a, err);
          }) : Promise.resolve();
        }catch(err){
          console.warn('[patch.relayout] unexpected error', err);
          return Promise.resolve();
        }
      };
      window.__PLOTLY_RELAYOUT_PATCHED__ = true;
      // Plotly.Lib.warn のうち "Relayout fail" を抑制（デバッグ時のみ）
      try{
        if(window.Plotly.Lib && typeof window.Plotly.Lib.warn === 'function' && !window.__PLOTLY_LIBWARN_PATCHED__){
          const origWarn = window.Plotly.Lib.warn;
          window.Plotly.Lib.warn = function patchedWarn(){
            try{
              const msg = arguments && arguments[0] ? String(arguments[0]) : '';
              if(msg.indexOf('Relayout fail') !== -1){
                // 2 つ目以降の引数（問題の updates など）を一応表示
                console.info('[plotly.warn suppressed] Relayout fail:', arguments[1], arguments[2]);
                return; // 抑制
              }
            }catch(_){/* noop */}
            return origWarn.apply(this, arguments);
          };
          window.__PLOTLY_LIBWARN_PATCHED__ = true;
          console.info('[patch.lib.warn] installed');
        }
      }catch(err){ console.warn('patch Plotly.Lib.warn failed', err); }

          // さらに保険として console.warn を薄くラップし、"Relayout fail" を info に降格（デバッグのみ）
          try{
            if(!window.__CONSOLE_WARN_PATCHED__ && typeof console !== 'undefined' && typeof console.warn === 'function'){
              const _origConsoleWarn = console.warn.bind(console);
              console.warn = function(){
                try{
                  const msg = arguments && arguments[0] ? String(arguments[0]) : '';
                  if(msg.indexOf('Relayout fail') !== -1){
                    console.info('[console.warn suppressed] Relayout fail:', arguments[1], arguments[2]);
                    return;
                  }
                }catch(_){/* noop */}
                return _origConsoleWarn.apply(console, arguments);
              };
              window.__CONSOLE_WARN_PATCHED__ = true;
              console.info('[patch.console.warn] installed');
            }
            // 同様に console.log/console.error でも出る可能性を抑止（デバッグのみ）
            if(!window.__CONSOLE_LOG_PATCHED__ && typeof console !== 'undefined' && typeof console.log === 'function'){
              const _origConsoleLog = console.log.bind(console);
              console.log = function(){
                try{
                  const msg = arguments && arguments[0] ? String(arguments[0]) : '';
                  if(msg.indexOf('Relayout fail') !== -1 || msg.indexOf('WARN: Relayout fail') !== -1){
                    console.info('[console.log suppressed] Relayout fail:', arguments[1], arguments[2]);
                    return;
                  }
                }catch(_){/* noop */}
                return _origConsoleLog.apply(console, arguments);
              };
              window.__CONSOLE_LOG_PATCHED__ = true;
              console.info('[patch.console.log] installed');
            }
            if(!window.__CONSOLE_ERROR_PATCHED__ && typeof console !== 'undefined' && typeof console.error === 'function'){
              const _origConsoleError = console.error.bind(console);
              console.error = function(){
                try{
                  const msg = arguments && arguments[0] ? String(arguments[0]) : '';
                  if(msg.indexOf('Relayout fail') !== -1){
                    console.info('[console.error suppressed] Relayout fail:', arguments[1], arguments[2]);
                    return;
                  }
                }catch(_){/* noop */}
                return _origConsoleError.apply(console, arguments);
              };
              window.__CONSOLE_ERROR_PATCHED__ = true;
              console.info('[patch.console.error] installed');
            }
          }catch(err){ /* ignore */ }
      // 追加: 内部 API Plots.relayout もパッチ（デバッグのみ）
      try{
        if(window.Plotly.Plots && typeof window.Plotly.Plots.relayout === 'function' && !window.__PLOTS_RELAYOUT_PATCHED__){
          const origPlotsRelayout = window.Plotly.Plots.relayout;
          let __relayoutFailSeq = 0;
          window.Plotly.Plots.relayout = function patchedPlotsRelayout(gd, a, b){
            try{
              if(typeof a === 'string'){
                if(typeof b === 'undefined'){
                  const id = ++__relayoutFailSeq;
                  console.groupCollapsed('[patch.plots.relayout]#'+id+' string key without value (ignore) key=', a);
                  console.trace('stack');
                  console.groupEnd();
                  return Promise.resolve();
                }
                const pr = origPlotsRelayout.call(window.Plotly.Plots, gd, a, b);
                return (pr && typeof pr.then === 'function') ? pr.catch(err => {
                  const id = ++__relayoutFailSeq;
                  console.groupCollapsed('[patch.plots.relayout]#'+id+' rejected (string key) key='+a);
                  console.log('value=', b);
                  console.log('error=', err);
                  console.trace('stack');
                  console.groupEnd();
                }) : Promise.resolve();
              }
              if(!a || typeof a !== 'object' || Array.isArray(a)){
                const id = ++__relayoutFailSeq;
                console.groupCollapsed('[patch.plots.relayout]#'+id+' invalid updates (expect plain object)');
                console.log('updates=', a);
                console.trace('stack');
                console.groupEnd();
                return Promise.resolve();
              }
              const pr = origPlotsRelayout.call(window.Plotly.Plots, gd, a);
              return (pr && typeof pr.then === 'function') ? pr.catch(err => {
                const id = ++__relayoutFailSeq;
                console.groupCollapsed('[patch.plots.relayout]#'+id+' rejected (object updates)');
                console.log('updates=', a);
                console.log('error=', err);
                console.trace('stack');
                console.groupEnd();
              }) : Promise.resolve();
            }catch(err){
              const id = ++__relayoutFailSeq;
              console.groupCollapsed('[patch.plots.relayout]#'+id+' unexpected error');
              console.log('error=', err);
              console.trace('stack');
              console.groupEnd();
              return Promise.resolve();
            }
          };
          window.__PLOTS_RELAYOUT_PATCHED__ = true;
          console.info('[patch.plots.relayout] installed');
        }
      }catch(err){ console.warn('patch Plots.relayout failed', err); }

      // 未処理拒否の既定ログを抑制（理由 undefined のみ／デバッグ時のみ）
      window.addEventListener('unhandledrejection', function(ev){
        try{
          if(!ev) return;
          // Plotly 由来かつ undefined 理由の拒否を抑制
          const r = ev.reason;
          const isUndefined = (typeof r === 'undefined');
          const srcMatch = (ev && ev.promise && typeof ev.promise === 'object');
          if(isUndefined){
            console.warn('[unhandledrejection] suppressed undefined reason from a promise (likely Plotly relayout)');
            ev.preventDefault();
          }
        }catch(_){/* noop */}
      });
      console.info('[patch.relayout] installed');
    }catch(err){ console.warn('patchGlobalRelayout failed', err); }
  })();

  // 包絡線範囲へフィット（初期描画・Autoscaleボタン・ダブルクリックで共通使用）
  function fitEnvelopeRanges(reason){
    try{
      if(!envelopeData || !envelopeData.length) return;
      const r = computeEnvelopeRanges(envelopeData);
      console.info('[Fit] 包絡線範囲へフィット:', reason || '');
      cachedEnvelopeRange = r; // キャッシュ更新
      safeRelayout(plotDiv, {
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
  // 自動解析: 旧ボタン無効化不要

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
  // 手動 processData 廃止（自動解析は processDataDirect 使用）

  // === Direct Input Processing ===
  function processDataDirect(){
    try{
      const L = parseFloat(wall_length_m.value);
      const alpha = parseFloat(alpha_factor.value);
      const c0 = parseFloat(c0_factor.value);
      const side = envelope_side.value;
      const specificDeformationValue = parseFloat(specific_deformation.value);
      const maxUltimateDeformationValue = parseFloat(max_ultimate_deformation.value);

      if(!isFinite(L) || !isFinite(alpha) || !isFinite(c0) || c0 < 0 || !isFinite(specificDeformationValue) || specificDeformationValue <= 0 || !isFinite(maxUltimateDeformationValue) || maxUltimateDeformationValue <= 0){
        console.warn('入力値が不正です');
        return;
      }
      
      // Calculate gamma_specific from 1/specificDeformationValue
      const gamma_specific = 1.0 / specificDeformationValue;
      
      // Calculate delta_u_max from 1/maxUltimateDeformationValue
      const delta_u_max = 1.0 / maxUltimateDeformationValue;

      // Generate envelope from direct input data
      envelopeData = generateEnvelope(rawData, side);
      if(envelopeData.length === 0){
        console.warn('包絡線の生成に失敗しました');
        return;
      }

      // Calculate characteristic points
      analysisResults = calculateJTCCMMetrics(envelopeData, gamma_specific, delta_u_max, L, alpha, c0);

      // Render results
      renderPlot(envelopeData, analysisResults);
      renderResults(analysisResults);

      if(downloadExcelButton) downloadExcelButton.disabled = false;
  if(generatePdfButton) generatePdfButton.disabled = false;
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
  function calculateJTCCMMetrics(envelope, gamma_specific, delta_u_max, L, alpha, c0){
    const results = {};

    // Determine the sign of the envelope (positive or negative side)
    const envelopeSign = envelope[0] && envelope[0].Load < 0 ? -1 : 1;

    // Find provisional global Pmax (used for yielding & Pu derivation; may lie after δu)
    const Pmax_global_pt = envelope.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), envelope[0]);
    const Pmax_global = Math.abs(Pmax_global_pt.Load);

    // Calculate Py using Line Method (Section III.1) with provisional global Pmax
    const Py_result = calculatePy_LineMethod(envelope, Pmax_global);
    results.Py = Py_result.Py;
    results.Py_gamma = Py_result.Py_gamma;
    results.lineI = Py_result.lineI;
    results.lineII = Py_result.lineII;
    results.lineIII = Py_result.lineIII;

  // Calculate Pu and μ using Perfect Elasto-Plastic Model (Section IV)
  const Pu_result = calculatePu_EnergyEquivalent(envelope, results.Py, Pmax_global, delta_u_max);
  Object.assign(results, Pu_result);

    // Override Pmax with value BEFORE ultimate displacement δu per user requirement
    const delta_u = results.delta_u; // from Pu_result
    if(isFinite(delta_u)){
      const prePts = envelope.filter(pt => Math.abs(pt.gamma) <= delta_u + 1e-12); // tolerance
      if(prePts.length){
        const Pmax_pre_pt = prePts.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), prePts[0]);
        results.Pmax_global = Pmax_global; // store original for reference
        results.Pmax = Math.abs(Pmax_pre_pt.Load);
        results.Pmax_gamma = Math.abs(Pmax_pre_pt.gamma);
      }else{
        // Fallback keep global
        results.Pmax_global = Pmax_global;
        results.Pmax = Pmax_global;
        results.Pmax_gamma = Math.abs(Pmax_global_pt.gamma);
      }
    }else{
      results.Pmax_global = Pmax_global;
      results.Pmax = Pmax_global;
      results.Pmax_gamma = Math.abs(Pmax_global_pt.gamma);
    }

    // === Second pass: Restrict Py (and dependent values) to pre-δu segment ===
    try{
      const du1 = results.delta_u;
      if(isFinite(du1) && du1 > 0){
        const envPre = envelope.filter(pt => Math.abs(pt.gamma) <= du1 + 1e-12);
        if(envPre.length >= 3){
          const pmaxPrePt = envPre.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), envPre[0]);
          const pmaxPre = Math.abs(pmaxPrePt.Load);
          // Recompute Py with pre-δu envelope
          const Py_pre = calculatePy_LineMethod(envPre, pmaxPre);
          results.Py = Py_pre.Py;
          results.Py_gamma = Py_pre.Py_gamma;
          results.lineI = Py_pre.lineI;
          results.lineII = Py_pre.lineII;
          results.lineIII = Py_pre.lineIII;

          // Recompute Pu/μ/δv/δu on pre-δu envelope with updated Py
          const Pu_pre = calculatePu_EnergyEquivalent(envPre, results.Py, pmaxPre, delta_u_max);
          Object.assign(results, Pu_pre);

          // Recompute Pmax with final δu restriction
          const du2 = results.delta_u;
          const envPre2 = envelope.filter(pt => Math.abs(pt.gamma) <= du2 + 1e-12);
          if(envPre2.length){
            const pmaxPrePt2 = envPre2.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), envPre2[0]);
            results.Pmax_global = Pmax_global;
            results.Pmax = Math.abs(pmaxPrePt2.Load);
            results.Pmax_gamma = Math.abs(pmaxPrePt2.gamma);
          }
        }
      }
    }catch(e){
      console.warn('Second-pass (pre-δu) Py/Pu 再計算に失敗:', e);
    }

    // Calculate P0 (Section V.1) using final results
    const P0_result = calculateP0(results, envelope, gamma_specific, c0);
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
  function calculatePu_EnergyEquivalent(envelope, Py, Pmax, delta_u_max){
    // Find δy (gamma where Load = Py on envelope)
    const pt_y = findPointAtLoad(envelope, Py);
    const delta_y = Math.abs(pt_y.gamma);

    // Initial stiffness K
    const K = Py / delta_y;

    // Find δu (Section IV.1 Step 9)
    const delta_u_candidate1 = findDeltaU_08Pmax(envelope, Pmax);
    const delta_u_candidate2 = delta_u_max; // rad from user input
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
  function calculateP0(results, envelope, gamma_specific, c0){
    const { Py, Pu, mu, Pmax } = results;

    // (a) Yield strength
    const p0_a = Py;

    // (b) Ductility-based
    let p0_b;
    const denom = 2 * mu - 1;
    if(denom > 0){
      // Pu / (1/ sqrt(2μ - 1)) * C0 = C0 * Pu * sqrt(2μ - 1)
      p0_b = c0 * Pu * Math.sqrt(denom);
    }else{
      // フォールバック（定義域外の場合は Py を採用）
      p0_b = Py;
    }

    // (c) Max strength
    const p0_c = Pmax * (2/3);

    // (d) Specific deformation - use gamma_specific directly
    const pt = findPointAtGamma(envelope, gamma_specific, 'gamma');
    const p0_d = pt ? Math.abs(pt.Load) : Pmax;

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

    // レンジの健全性チェック（NaN / Infinity を排除）
    function sanitizeRange(arr, defMin, defMax){
      const a0 = Array.isArray(arr) ? arr[0] : undefined;
      const a1 = Array.isArray(arr) ? arr[1] : undefined;
      let minV = Number.isFinite(a0) ? a0 : defMin;
      let maxV = Number.isFinite(a1) ? a1 : defMax;
      if(!Number.isFinite(minV)) minV = -1;
      if(!Number.isFinite(maxV)) maxV = 1;
      if(minV === maxV){ maxV = minV + 1; }
      if(minV > maxV){ const t = minV; minV = maxV; maxV = t; }
      return [minV, maxV];
    }
    const xRangeSafe = sanitizeRange(ranges && ranges.xRange, -1, 1);
    const yRangeSafe = sanitizeRange(ranges && ranges.yRange, -1, 1);
    
    // 包絡線データを編集可能にするための状態管理
    let editableEnvelope = envelope.map(pt => ({...pt}));    // Original raw data (all points) - showing positive and negative loads
    const trace_rawdata = {
      x: rawData.map(pt => pt.gamma), // rad
      y: rawData.map(pt => pt.Load), // Keep original sign
      mode: 'lines+markers',
      name: '実験データ',
      line: {color: 'lightblue', width: 1},
      marker: {color: 'lightblue', size: 4},
      hoverinfo: 'skip' // 試験データのホバーを無効化し、包絡線点を優先
    };

    // Envelope line - keep original sign
    const trace_env = {
      x: editableEnvelope.map(pt => pt.gamma),
      y: editableEnvelope.map(pt => pt.Load), // Keep original sign from filtered data
      mode: 'lines',
      name: '包絡線',
      line: {color: 'blue', width: 2},
      hoverinfo: 'skip' // 包絡線ラインのホバーを無効化
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
  let gamma_max = Math.max(...envelope.map(pt => Math.abs(pt.gamma)));
    if(!Number.isFinite(gamma_max) || gamma_max <= 0){
      // 範囲から安全な最大を推定
      gamma_max = Math.max(Math.abs(xRangeSafe[0]), Math.abs(xRangeSafe[1]));
      if(!Number.isFinite(gamma_max) || gamma_max <= 0) gamma_max = 1;
    }
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
        range: xRangeSafe,
        autorange: false
      },
      yaxis: {
        title: '荷重 P (kN)',
        range: yRangeSafe,
        autorange: false
      },
      hovermode: 'closest',
      dragmode: 'pan', // デフォルトはパン操作
      showlegend: true,
      height: 600,
      uirevision: 'fixed',
      annotations: [
        // 終局変位 δu (rad) → Line VI の終点（delta_u の位置）に表示
        {
          x: (lineVI.gamma_end) * envelopeSign,
          y: (lineVI.Load) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `δu=${formatReciprocal(delta_u)}`,
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
          text: `Py=${Py.toFixed(1)} kN\nδy=${formatReciprocal(Py_gamma)}`,
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
          text: `δv=${formatReciprocal(delta_v)}`,
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
    // Box select / Lasso select を有効化
    displaylogo: false,
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
                safeRelayout(plotDiv, {
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
    })
    .catch(function(err){
      // 初期 newPlot の undefined 拒否などを握りつぶしてノイズ抑制
      console.info('[plot.init suppressed]', err);
    });
  }

  // === 包絡線点の編集機能 ===
  function setupEnvelopeEditing(editableEnvelope){
    let isDragging = false;
    let dragPointIndex = -1;
    let selectedPointIndex = -1; // Del キー用の選択状態
    let selectedPoints = []; // Box/Lasso select用の複数選択状態
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
    
    // Box Select / Lasso Select で複数点選択を検出
    plotDiv.on('plotly_selected', function(eventData){
      if(!eventData || !eventData.points) {
        selectedPoints = [];
        return;
      }
      // 包絡線点トレース（curveNumber === 2）のみを抽出
      selectedPoints = eventData.points
        .filter(pt => pt.curveNumber === 2)
        .map(pt => pt.pointIndex);
      console.debug('[plotly_selected] 選択された包絡線点:', selectedPoints);
    });
    
    // 選択解除
    plotDiv.on('plotly_deselect', function(){
      selectedPoints = [];
      console.debug('[plotly_deselect] 選択解除');
    });
    
    // ダブルクリックによる点追加やデフォルト操作は許容（別処理はしない）
    
    // ダブルクリックによる新規点追加機能は廃止（仕様変更）
    
    // Delキーで選択中の点を削除
    // 既存のキーリスナーを解除
    if(_keydownHandler){
      document.removeEventListener('keydown', _keydownHandler);
    }
    const handleKeydown = function(e){
      // Delキーで選択中の点を削除（単一点 or 複数点）
      if(e.key === 'Delete' || e.key === 'Del'){
        // Box/Lasso selectで複数選択がある場合
        if(selectedPoints.length > 0){
          // 最小2点は残すためのチェック
          const remainingCount = editableEnvelope.length - selectedPoints.length;
          if(remainingCount < 2){
            alert('包絡線には最低2点が必要です。削除できません。');
            return;
          }
          // 降順ソートして後ろから削除（インデックスずれ防止）
          const sortedIndices = [...selectedPoints].sort((a,b) => b - a);
          sortedIndices.forEach(idx => {
            if(idx >= 0 && idx < editableEnvelope.length){
              editableEnvelope.splice(idx, 1);
            }
          });
          pushHistory(editableEnvelope);
          appendLog(`包絡線点 ${selectedPoints.length}個 を一括削除しました`);
          selectedPoints = [];
          selectedPointIndex = -1;
          window._selectedEnvelopePoint = -1;
          recalculateFromEnvelope(editableEnvelope);
          return;
        }
        // 単一点選択の場合（既存の動作）
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
    
    // Shift + ドラッグで包絡線点の座標を変更
    let shiftDragging = false;
    let shiftDragIndex = -1;
    let shiftDragStartX = 0;
    let shiftDragStartY = 0;
    
    plotDiv.on('plotly_hover', function(data){
      if(!data.points || data.points.length === 0) return;
      const pt = data.points[0];
      // 包絡線点（curveNumber === 2）にホバー時、カーソルをポインタに
      if(pt.curveNumber === 2){
        plotDiv.style.cursor = 'pointer';
      } else {
        plotDiv.style.cursor = 'default';
      }
    });
    
    plotDiv.on('plotly_unhover', function(){
      if(!shiftDragging){
        plotDiv.style.cursor = 'default';
      }
    });
    
    // マウスダウン: Ctrl+Shift 押下中かつ包絡線点上ならドラッグ開始
    let mousedownHandler = function(e){
      // Ctrl+Shift 同時押しでのみドラッグ開始（Mac の command 対応は ctrlKey 優先）
      const ctrlOrMeta = e.ctrlKey || e.metaKey;
      if(!(ctrlOrMeta && e.shiftKey)) return;
      
      // 選択中の点が存在しない場合はドラッグ不可
      if(window._selectedEnvelopePoint < 0 || !editableEnvelope || window._selectedEnvelopePoint >= editableEnvelope.length){
        return;
      }
      
      // Plotlyのイベントから選択点の座標を取得
      const xaxis = plotDiv._fullLayout.xaxis;
      const yaxis = plotDiv._fullLayout.yaxis;
      if(!xaxis || !yaxis) return;
      
      const bbox = (dragLayer || plotDiv).getBoundingClientRect();
      const clickX = e.clientX - bbox.left;
      const clickY = e.clientY - bbox.top;
      
      // 選択点のピクセル座標を計算
      const selectedIdx = window._selectedEnvelopePoint;
      const selectedPt = editableEnvelope[selectedIdx];
      const px = xaxis.l2p(selectedPt.gamma);
      const py = yaxis.l2p(selectedPt.Load);
      const dist = Math.sqrt((clickX - px)**2 + (clickY - py)**2);
      
      // 35px以内なら選択点と判定（ヒット領域）
      if(dist < 35){
        // Plotlyのデフォルトドラッグを無効化（先に実行）
        e.stopImmediatePropagation();
        e.preventDefault();
        
        shiftDragging = true;
        shiftDragIndex = selectedIdx;
        shiftDragStartX = clickX;
        shiftDragStartY = clickY;
        plotDiv.style.cursor = 'move';
        
        // Plotlyデフォルトのズーム/パンはイベント抑止で無効化（dragmodeの変更は行わない）
        
        // ツールチップ表示
        if(pointTooltip){
          pointTooltip.textContent = `γ: ${selectedPt.gamma.toFixed(6)}, P: ${selectedPt.Load.toFixed(3)}`;
          pointTooltip.style.left = e.clientX + 10 + 'px';
          pointTooltip.style.top = e.clientY + 10 + 'px';
          pointTooltip.style.display = 'block';
        }
      }
    };
    
    // マウスムーブ: ドラッグ中なら座標更新
    let mousemoveHandler = function(e){
      if(!shiftDragging || shiftDragIndex < 0) return;
      
      const xaxis = plotDiv._fullLayout.xaxis;
      const yaxis = plotDiv._fullLayout.yaxis;
      if(!xaxis || !yaxis) return;
      
      const bbox = (dragLayer || plotDiv).getBoundingClientRect();
      const moveX = e.clientX - bbox.left;
      const moveY = e.clientY - bbox.top;
      
  // データ座標に変換（pixel → linear）
  const newGamma = xaxis.p2l(moveX);
  const newLoad = yaxis.p2l(moveY);
      
      // 包絡線点を更新
      editableEnvelope[shiftDragIndex].gamma = newGamma;
      editableEnvelope[shiftDragIndex].Load = newLoad;
      
      // プロット更新
      updateEnvelopePlot(editableEnvelope);
      
      // 点編集ダイアログが対象点を編集中なら、入力欄をリアルタイム更新（ユーザーのリクエスト対応）
      if(pointEditDialog && pointEditDialog.style.display !== 'none' && window._selectedEnvelopePoint === shiftDragIndex){
        // 表示精度は既存UIに合わせる（γ: 4～6桁, P: 3桁）必要に応じて後で統一可能
        if(editGammaInput){ editGammaInput.value = newGamma.toFixed(4); }
        if(editLoadInput){ editLoadInput.value = newLoad.toFixed(1); }
      }

      // ツールチップ更新
      if(pointTooltip){
        pointTooltip.textContent = `γ: ${newGamma.toFixed(6)}, P: ${newLoad.toFixed(3)}`;
        pointTooltip.style.left = e.clientX + 10 + 'px';
        pointTooltip.style.top = e.clientY + 10 + 'px';
      }
      
      e.preventDefault();
    };
    
    // マウスアップ: ドラッグ終了
    let mouseupHandler = function(e){
      if(shiftDragging && shiftDragIndex >= 0){
        // 履歴に保存
        pushHistory(editableEnvelope);
        
        // 再計算
        recalculateFromEnvelope(editableEnvelope);
        appendLog(`包絡線点 #${shiftDragIndex} をドラッグ移動しました`);
        
        // ツールチップ非表示
        if(pointTooltip){
          pointTooltip.style.display = 'none';
        }
        
        shiftDragging = false;
        shiftDragIndex = -1;
        plotDiv.style.cursor = 'default';
        
        // Plotlyのドラッグモードを復元（エラーは握りつぶし）
        if(window.Plotly && plotDiv){
          try{
            safeRelayout(plotDiv, {'dragmode': 'pan'});
          }catch(_){/* noop */}
        }
      }
    };
    
    // イベントリスナー登録（キャプチャフェーズで先にイベントを取得）
    const dragLayer = plotDiv.querySelector('.draglayer') || plotDiv;
    // マウスイベント
    dragLayer.addEventListener('mousedown', mousedownHandler, true);
    document.addEventListener('mousemove', mousemoveHandler);
    document.addEventListener('mouseup', mouseupHandler);
    document.addEventListener('mouseleave', mouseupHandler);
    // ポインタイベント（Zoom優先を抑止するため早期にハンドリング）
    const pointerdownHandler = function(e){
      if(e && (e.pointerType === 'mouse' || e.pointerType === 'pen')){
        mousedownHandler(e);
      }
    };
    const pointermoveHandler = function(e){
      if(e && (e.pointerType === 'mouse' || e.pointerType === 'pen')){
        mousemoveHandler(e);
      }
    };
    const pointerupHandler = function(e){
      if(e && (e.pointerType === 'mouse' || e.pointerType === 'pen')){
        mouseupHandler(e);
      }
    };
    dragLayer.addEventListener('pointerdown', pointerdownHandler, true);
    document.addEventListener('pointermove', pointermoveHandler, true);
    document.addEventListener('pointerup', pointerupHandler, true);
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
      
  const x1 = xaxis.l2p(p1.gamma);
  const y1 = yaxis.l2p(p1.Load);
  const x2 = xaxis.l2p(p2.gamma);
  const y2 = yaxis.l2p(p2.Load);
      
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
      const c0 = parseFloat(c0_factor.value);
      const specificDeformationValue = parseFloat(specific_deformation.value);
      const maxUltimateDeformationValue = parseFloat(max_ultimate_deformation.value);
      
      if(!isFinite(L) || !isFinite(alpha) || !isFinite(c0) || c0 < 0 || !isFinite(specificDeformationValue) || specificDeformationValue <= 0 || !isFinite(maxUltimateDeformationValue) || maxUltimateDeformationValue <= 0) return;
      
      const gamma_specific = 1.0 / specificDeformationValue;
      const delta_u_max = 1.0 / maxUltimateDeformationValue;
      
      analysisResults = calculateJTCCMMetrics(envelopeData, gamma_specific, delta_u_max, L, alpha, c0);
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
    document.getElementById('val_dy').textContent = formatReciprocal(r.delta_y);
    document.getElementById('val_K').textContent = r.K.toFixed(2);
    document.getElementById('val_pu').textContent = r.Pu.toFixed(3);
    document.getElementById('val_dv').textContent = formatReciprocal(r.delta_v);
    document.getElementById('val_du').textContent = formatReciprocal(r.delta_u);
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
    // 1m当たり (kN/m)
    const Lval = parseFloat(wall_length_m.value);
    const paPerEl = document.getElementById('val_pa_per_m');
    const puPerEl = document.getElementById('val_pu_per_m');
    if(isFinite(Lval) && Lval>0){
      if(paPerEl) paPerEl.textContent = (r.Pa / Lval).toFixed(3);
      if(puPerEl) puPerEl.textContent = (r.Pu / Lval).toFixed(3);
    } else {
      if(paPerEl) paPerEl.textContent = '-';
      if(puPerEl) puPerEl.textContent = '-';
    }
    document.getElementById('val_magnification').textContent = r.magnification_rounded.toFixed(1) + ' 倍';
  }

  async function downloadExcel(){
    if(!window.ExcelJS){
      alert('ExcelJSライブラリが読み込まれていません');
      return;
    }
    try{
      const specimen = (specimen_name && specimen_name.value ? specimen_name.value.trim() : 'testname');
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
      wsSummary.addRow(['試験体名称', specimen, '']);
      const Lval2 = parseFloat(wall_length_m.value);
      const rows = [
        ['最大耐力 Pmax', r.Pmax, 'kN'],
        ['降伏耐力 Py', r.Py, 'kN'],
        ['降伏変位 δy', formatReciprocal(r.delta_y), '1/n'],
        ['初期剛性 K', r.K, 'kN/rad'],
        ['終局耐力 Pu', r.Pu, 'kN'],
        ['終局変位 δu', formatReciprocal(r.delta_u), '1/n'],
        ['塑性率 μ', r.mu, ''],
        ['P0(a) 降伏耐力', r.p0_a, 'kN'],
        ['P0(b) 靭性基準', r.p0_b, 'kN'],
        ['P0(c) 最大耐力基準', r.p0_c, 'kN'],
        ['P0(d) 特定変形時', r.p0_d, 'kN'],
        ['短期基準せん断耐力 P0', r.P0, 'kN'],
        ['短期許容せん断耐力 Pa', r.Pa, 'kN'],
        ['短期許容せん断耐力 Pa (kN/m)', (isFinite(Lval2)&&Lval2>0)? r.Pa/Lval2 : '-', 'kN/m'],
        ['終局耐力 Pu (kN/m)', (isFinite(Lval2)&&Lval2>0)? r.Pu/Lval2 : '-', 'kN/m'],
        ['壁倍率', r.magnification_rounded, '倍']
      ];
      rows.forEach(row => wsSummary.addRow(row));
      wsSummary.columns.forEach(col => { col.width = 22; });
      // 数値フォーマット適用
      for(let i=2;i<=wsSummary.rowCount;i++){
        const label = wsSummary.getCell(i,1).value;
        const cell = wsSummary.getCell(i,2);
        if(typeof cell.value !== 'number') continue;
        if(wsSummary.getCell(i,3).value === '1/n') continue; // reciprocalは文字列のまま
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
  const excelFileName = `Results_${specimen.replace(/[^a-zA-Z0-9_\-一-龥ぁ-んァ-ヶ]/g,'_')}.xlsx`;
  a.download = excelFileName;
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
})();
