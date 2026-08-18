/**
 * IT Asset Sticker Studio - Main Application Logic
 */

document.addEventListener('DOMContentLoaded', () => {
  
  // ==========================================
  // Sample Data (Based on User Screenshot)
  // ==========================================
  const SAMPLE_DATA = [
    { "SN": 165, "Hostname": "RFCPL-LT-004", "Serial": "9CFLC44", "Employee Name": "Akhil Balakrishnan", "Department": "Treasury" },
    { "SN": 268, "Hostname": "RFCPL-LT-406", "Serial": "BM5NQ74", "Employee Name": "Nikhil Sandav", "Department": "Treasury" },
    { "SN": 405, "Hostname": "RFCPL-LT-548", "Serial": "LQ261399NM", "Employee Name": "Anshuman", "Department": "Technology" },
    { "SN": 220, "Hostname": "RFCPL-LT-0360", "Serial": "CQYY792V3X", "Employee Name": "Ashutosh Mishra", "Department": "Technology" },
    { "SN": 332, "Hostname": "RFCPL-LT-472", "Serial": "MJ4V6LF40N", "Employee Name": "Bharat Rathore", "Department": "Technology" },
    { "SN": 455, "Hostname": "RFCPL-LT-601", "Serial": "DX005ZJL", "Employee Name": "Bhavya Momaya", "Department": "Technology" },
    { "SN": 520, "Hostname": "RFCPL-LT-219", "Serial": "PG02XXTD", "Employee Name": "danish shaikh", "Department": "Technology" },
    { "SN": 155, "Hostname": "RFCPL-LT-170", "Serial": "DGM2B24", "Employee Name": "Devanand Giri", "Department": "Technology" },
    { "SN": 262, "Hostname": "RFCPL-MAC-400", "Serial": "KM96951216", "Employee Name": "Divit Rao", "Department": "Technology" },
    { "SN": 199, "Hostname": "RFCPL-LT-093", "Serial": "L9YR7HLX22", "Employee Name": "Gourav Dangi", "Department": "Technology" },
    { "SN": 243, "Hostname": "RFCPL-LT-399", "Serial": "PF58Y3DA", "Employee Name": "Harshal dalvi", "Department": "Technology" },
    { "SN": 581, "Hostname": "RFCPL-LT-358", "Serial": "M5YWXKR2JV", "Employee Name": "Kamalkant Maharana", "Department": "Technology" },
    { "SN": 411, "Hostname": "RFCPL-LT-557", "Serial": "L6D3HF161H", "Employee Name": "Karthik D", "Department": "Technology" },
    { "SN": 304, "Hostname": "RFCPL-LT-469", "Serial": "3D3BR24", "Employee Name": "Mohammed Ussaid", "Department": "Technology" },
    { "SN": 76,  "Hostname": "RFCPL-LT-229", "Serial": "GLHSFX3", "Employee Name": "Nikhil Kothakota", "Department": "Technology" },
    { "SN": 122, "Hostname": "RFCPL-LT-141", "Serial": "3BM2B24", "Employee Name": "Nikhilesh Thakur", "Department": "Technology" },
    { "SN": 241, "Hostname": "RFCPL-LT-392", "Serial": "J73TL00XHW", "Employee Name": "Nilesh Joge", "Department": "Technology" },
    { "SN": 88,  "Hostname": "RFCPL-LT-098", "Serial": "HRW17J9FG4", "Employee Name": "Pooja Patel", "Department": "Technology" },
    { "SN": 263, "Hostname": "RFCPL-LT-401", "Serial": "JPX21C799H", "Employee Name": "Pratish Srivastava", "Department": "Technology" },
    { "SN": 77,  "Hostname": "RFCPL-LT-232", "Serial": "PF3B2M53", "Employee Name": "Prince Kumar", "Department": "Technology" },
    { "SN": 264, "Hostname": "RFCPL-MAC-402", "Serial": "KHQH701Q0J", "Employee Name": "Raj Shetye", "Department": "Technology" }
  ];

  // ==========================================
  // Application State
  // ==========================================
  const state = {
    dataset: [],
    columns: [],
    mappings: {
      host: '',
      asset: '',
      serial: ''
    },
    config: {
      colGap: 2,        // mm (2 mm gap)
      rowGap: 4,        // mm (4 mm gap)
      refGap: 3,        // px (3 px space above block)
      gridCols: 4,      // 4 columns per row
      blockWidth: 44,   // mm (44 mm block width)
      blockHeight: 13,  // mm (13 mm block height)
      blockPadding: 2,  // px (2 px block padding)
      fontHost: 7,      // px
      fontAsset: 14,    // px
      fontSerial: 11,   // px
      textAlign: 'center',
      prefixHost: 'Host:',
      prefixSerial: 'S/N:',
      borderStyle: 'rounded',
      barcodeToggle: 'none',
      pageSize: 'A4',
      pageMargin: 10    // mm
    },
    zoom: 100,
    showCutMarks: true
  };

  // DOM Elements References
  const dropzone = document.getElementById('excel-dropzone');
  const fileInput = document.getElementById('excel-file-input');
  const stickerGrid = document.getElementById('sticker-grid');
  const paperSheet = document.getElementById('paper-sheet');
  const paperViewport = document.getElementById('paper-viewport');
  const dataCountBadge = document.getElementById('data-count-badge');
  const countSummary = document.getElementById('count-summary');

  // Mapping Selects
  const mapHostSelect = document.getElementById('map-host');
  const mapAssetSelect = document.getElementById('map-asset');
  const mapSerialSelect = document.getElementById('map-serial');

  // Controls Elements
  const ctrlColGap = document.getElementById('ctrl-col-gap');
  const ctrlRowGap = document.getElementById('ctrl-row-gap');
  const ctrlRefGap = document.getElementById('ctrl-ref-gap');
  const ctrlGridCols = document.getElementById('ctrl-grid-cols');
  const ctrlBlockWidth = document.getElementById('ctrl-block-width');
  const ctrlBlockHeight = document.getElementById('ctrl-block-height');
  const ctrlBlockPadding = document.getElementById('ctrl-block-padding');

  const ctrlFontHost = document.getElementById('ctrl-font-host');
  const ctrlFontAsset = document.getElementById('ctrl-font-asset');
  const ctrlFontSerial = document.getElementById('ctrl-font-serial');

  const ctrlPrefixHost = document.getElementById('ctrl-prefix-host');
  const ctrlPrefixSerial = document.getElementById('ctrl-prefix-serial');
  const ctrlBorderStyle = document.getElementById('ctrl-border-style');
  const ctrlBarcodeToggle = document.getElementById('ctrl-barcode-toggle');
  const ctrlPageSize = document.getElementById('ctrl-page-size');
  const ctrlPageMargin = document.getElementById('ctrl-page-margin');

  // Value Badges
  const valColGap = document.getElementById('val-col-gap');
  const valRowGap = document.getElementById('val-row-gap');
  const valRefGap = document.getElementById('val-ref-gap');
  const valGridCols = document.getElementById('val-grid-cols');
  const valBlockWidth = document.getElementById('val-block-width');
  const valBlockHeight = document.getElementById('val-block-height');
  const valBlockPadding = document.getElementById('val-block-padding');
  const valFontHost = document.getElementById('val-font-host');
  const valFontAsset = document.getElementById('val-font-asset');
  const valFontSerial = document.getElementById('val-font-serial');
  const valPageMargin = document.getElementById('val-page-margin');

  // ==========================================
  // Initialization & Event Binding
  // ==========================================
  function init() {
    bindControlEvents();
    bindDropzoneEvents();
    bindModalEvents();
    bindZoomEvents();
    
    // Load sample data by default so user immediately sees results!
    loadDataset(SAMPLE_DATA);
  }

  // Load dataset into state & update selectors
  function loadDataset(data) {
    if (!data || data.length === 0) return;
    state.dataset = data;
    state.columns = Object.keys(data[0]);

    // Populate Column Selectors
    populateMappingSelects();

    // Auto-detect columns
    autoDetectColumns();

    // Update UI Badges & Render
    dataCountBadge.textContent = data.length;
    countSummary.textContent = `Showing ${data.length} stickers`;

    renderTable();
    renderStickers();
  }

  // Populate drop-down options for column mapping
  function populateMappingSelects() {
    const selects = [mapHostSelect, mapAssetSelect, mapSerialSelect];
    selects.forEach(select => {
      select.innerHTML = '<option value="">-- Select Column --</option>';
      state.columns.forEach(col => {
        const opt = document.createElement('option');
        opt.value = col;
        opt.textContent = col;
        select.appendChild(opt);
      });
    });
  }

  // Smart auto-detection of column headers
  function autoDetectColumns() {
    state.columns.forEach(col => {
      const lower = col.toLowerCase().trim();
      if (lower.includes('employee') || lower.includes('user') || lower.includes('assigned') || (lower.includes('host') && lower.includes('name'))) {
        state.mappings.host = col;
        mapHostSelect.value = col;
      } else if (lower === 'hostname' || lower.includes('asset') || lower.includes('tag') || lower.includes('code')) {
        state.mappings.asset = col;
        mapAssetSelect.value = col;
      } else if (lower.includes('serial') || lower.includes('sn') || lower.includes('s/n') || lower.includes('service')) {
        state.mappings.serial = col;
        mapSerialSelect.value = col;
      }
    });

    // Fallback if not matched
    if (!state.mappings.host && state.columns[3]) {
      state.mappings.host = state.columns[3];
      mapHostSelect.value = state.columns[3];
    }
    if (!state.mappings.asset && state.columns[1]) {
      state.mappings.asset = state.columns[1];
      mapAssetSelect.value = state.columns[1];
    }
    if (!state.mappings.serial && state.columns[2]) {
      state.mappings.serial = state.columns[2];
      mapSerialSelect.value = state.columns[2];
    }
  }

  // ==========================================
  // Sticker Render Engine
  // ==========================================
  function renderStickers() {
    stickerGrid.innerHTML = '';

    // Apply grid gap & column properties
    stickerGrid.style.display = 'grid';
    stickerGrid.style.gridTemplateColumns = `repeat(${state.config.gridCols}, 1fr)`;
    stickerGrid.style.columnGap = `${state.config.colGap}mm`;
    stickerGrid.style.rowGap = `${state.config.rowGap}mm`;

    // Apply paper sheet page styling
    paperSheet.style.padding = `${state.config.pageMargin}mm`;
    if (state.config.pageSize === 'A4') {
      paperSheet.style.width = '210mm';
      paperSheet.style.minHeight = '297mm';
    } else if (state.config.pageSize === 'Letter') {
      paperSheet.style.width = '8.5in';
      paperSheet.style.minHeight = '11in';
    } else {
      paperSheet.style.width = 'auto';
      paperSheet.style.minHeight = 'auto';
    }

    if (state.showCutMarks) {
      paperSheet.classList.add('show-cut-marks');
    } else {
      paperSheet.classList.remove('show-cut-marks');
    }

    // Iterate dataset and construct HTML elements
    state.dataset.forEach((row, index) => {
      const hostVal = state.mappings.host ? (row[state.mappings.host] || '') : '';
      const assetVal = state.mappings.asset ? (row[state.mappings.asset] || '') : '';
      const serialVal = state.mappings.serial ? (row[state.mappings.serial] || '') : '';

      const pHost = state.config.prefixHost ? (state.config.prefixHost.endsWith(' ') ? state.config.prefixHost : state.config.prefixHost + ' ') : '';
      const pSerial = state.config.prefixSerial ? (state.config.prefixSerial.endsWith(' ') ? state.config.prefixSerial : state.config.prefixSerial + ' ') : '';

      // Outer Wrapper
      const itemWrapper = document.createElement('div');
      itemWrapper.className = 'sticker-item-wrapper';

      // 1. Outside & Above: Host Name Reference
      const refHostEl = document.createElement('div');
      refHostEl.className = 'sticker-ref-host';
      refHostEl.style.fontSize = `${state.config.fontHost}px`;
      refHostEl.style.marginBottom = `${state.config.refGap}px`;
      refHostEl.textContent = pHost + hostVal;

      // 2. Sticker Block (Bordered Box)
      const blockEl = document.createElement('div');
      blockEl.className = `sticker-block border-${state.config.borderStyle}`;
      blockEl.style.width = `${state.config.blockWidth}mm`;
      blockEl.style.height = `${state.config.blockHeight}mm`;
      blockEl.style.padding = `${state.config.blockPadding}px`;
      blockEl.style.textAlign = state.config.textAlign;

      // Asset Tag
      const assetEl = document.createElement('div');
      assetEl.className = 'sticker-asset-tag';
      assetEl.style.fontSize = `${state.config.fontAsset}px`;
      assetEl.textContent = assetVal;

      // Serial Number
      const serialEl = document.createElement('div');
      serialEl.className = 'sticker-serial-number';
      serialEl.style.fontSize = `${state.config.fontSerial}px`;
      serialEl.textContent = pSerial + serialVal;

      blockEl.appendChild(assetEl);
      blockEl.appendChild(serialEl);

      // Optional Barcode Generation
      if (state.config.barcodeToggle !== 'none' && window.JsBarcode) {
        const barcodeVal = (state.config.barcodeToggle === 'asset') ? assetVal : serialVal;
        if (barcodeVal) {
          const svgEl = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          svgEl.className = 'sticker-barcode-svg';
          try {
            JsBarcode(svgEl, barcodeVal, {
              format: 'CODE128',
              displayValue: false,
              height: 25,
              margin: 2
            });
            blockEl.appendChild(svgEl);
          } catch (e) {
            console.warn('Barcode render error:', e);
          }
        }
      }

      itemWrapper.appendChild(refHostEl);
      itemWrapper.appendChild(blockEl);
      stickerGrid.appendChild(itemWrapper);
    });
  }

  // ==========================================
  // Controls Event Listeners
  // ==========================================
  function bindControlEvents() {
    // Accordion Toggle Listeners
    document.querySelectorAll('.accordion-header').forEach(header => {
      header.addEventListener('click', (e) => {
        const section = header.parentElement;
        section.classList.toggle('active');
      });
    });

    // Mobile / Responsive Sidebar Toggle
    const sidebar = document.getElementById('app-sidebar');
    const toggleBtn = document.getElementById('btn-toggle-sidebar');
    if (toggleBtn && sidebar) {
      toggleBtn.addEventListener('click', () => {
        sidebar.classList.toggle('open');
      });
    }

    // Column Mapping Selects
    mapHostSelect.addEventListener('change', (e) => {
      state.mappings.host = e.target.value;
      renderStickers();
    });
    mapAssetSelect.addEventListener('change', (e) => {
      state.mappings.asset = e.target.value;
      renderStickers();
    });
    mapSerialSelect.addEventListener('change', (e) => {
      state.mappings.serial = e.target.value;
      renderStickers();
    });

    // Spacing Sliders
    ctrlColGap.addEventListener('input', (e) => {
      state.config.colGap = parseInt(e.target.value);
      valColGap.textContent = `${state.config.colGap} mm`;
      renderStickers();
    });

    ctrlRowGap.addEventListener('input', (e) => {
      state.config.rowGap = parseInt(e.target.value);
      valRowGap.textContent = `${state.config.rowGap} mm`;
      renderStickers();
    });

    ctrlRefGap.addEventListener('input', (e) => {
      state.config.refGap = parseInt(e.target.value);
      valRefGap.textContent = `${state.config.refGap} px`;
      renderStickers();
    });

    ctrlGridCols.addEventListener('input', (e) => {
      state.config.gridCols = parseInt(e.target.value);
      valGridCols.textContent = state.config.gridCols;
      renderStickers();
    });

    ctrlBlockWidth.addEventListener('input', (e) => {
      state.config.blockWidth = parseInt(e.target.value);
      valBlockWidth.textContent = `${state.config.blockWidth} mm`;
      renderStickers();
    });

    ctrlBlockHeight.addEventListener('input', (e) => {
      state.config.blockHeight = parseInt(e.target.value);
      valBlockHeight.textContent = `${state.config.blockHeight} mm`;
      renderStickers();
    });

    ctrlBlockPadding.addEventListener('input', (e) => {
      state.config.blockPadding = parseInt(e.target.value);
      valBlockPadding.textContent = `${state.config.blockPadding} px`;
      renderStickers();
    });

    // Typography Sliders
    ctrlFontHost.addEventListener('input', (e) => {
      state.config.fontHost = parseInt(e.target.value);
      valFontHost.textContent = `${state.config.fontHost} px`;
      renderStickers();
    });

    ctrlFontAsset.addEventListener('input', (e) => {
      state.config.fontAsset = parseInt(e.target.value);
      valFontAsset.textContent = `${state.config.fontAsset} px`;
      renderStickers();
    });

    ctrlFontSerial.addEventListener('input', (e) => {
      state.config.fontSerial = parseInt(e.target.value);
      valFontSerial.textContent = `${state.config.fontSerial} px`;
      renderStickers();
    });

    // Align Toggle Group
    const alignBtns = document.querySelectorAll('#align-toggle-group .btn-toggle');
    alignBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        alignBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.config.textAlign = btn.dataset.align;
        renderStickers();
      });
    });

    // Prefixes & Formatting
    ctrlPrefixHost.addEventListener('input', (e) => {
      state.config.prefixHost = e.target.value;
      renderStickers();
    });
    ctrlPrefixSerial.addEventListener('input', (e) => {
      state.config.prefixSerial = e.target.value;
      renderStickers();
    });
    ctrlBorderStyle.addEventListener('change', (e) => {
      state.config.borderStyle = e.target.value;
      renderStickers();
    });
    ctrlBarcodeToggle.addEventListener('change', (e) => {
      state.config.barcodeToggle = e.target.value;
      renderStickers();
    });

    // Page Setup
    ctrlPageSize.addEventListener('change', (e) => {
      state.config.pageSize = e.target.value;
      renderStickers();
    });
    ctrlPageMargin.addEventListener('input', (e) => {
      state.config.pageMargin = parseInt(e.target.value);
      valPageMargin.textContent = `${state.config.pageMargin} mm`;
      renderStickers();
    });

    // Action Header Buttons
    document.getElementById('btn-load-sample').addEventListener('click', () => {
      loadDataset(SAMPLE_DATA);
    });

    document.getElementById('btn-print').addEventListener('click', () => {
      window.print();
    });

    document.getElementById('toggle-cut-marks').addEventListener('change', (e) => {
      state.showCutMarks = e.target.checked;
      renderStickers();
    });
  }

  // ==========================================
  // File Upload & Excel Parsing (SheetJS)
  // ==========================================
  function bindDropzoneEvents() {
    dropzone.addEventListener('click', () => fileInput.click());

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) {
        handleFileSelect(e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
      }
    });
  }

  function handleFileSelect(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const parsedJson = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        if (parsedJson && parsedJson.length > 0) {
          loadDataset(parsedJson);
        } else {
          alert('Uploaded sheet appears to be empty.');
        }
      } catch (err) {
        console.error('Error parsing Excel file:', err);
        alert('Could not parse Excel file. Please ensure it is a valid .xlsx, .xls, or .csv file.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // Helper for XML escaping
  function escapeXML(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  // ==========================================
  // SVG Vector Generator Engine
  // ==========================================
  function generateFullSheetSVG() {
    if (!state.dataset || state.dataset.length === 0) return '';

    const MM_TO_PX = 3.7795275591; // 96 DPI conversion

    // Config parameters
    const pageMarginMm = state.config.pageMargin;
    const gridCols = state.config.gridCols;
    const colGapMm = state.config.colGap;
    const rowGapMm = state.config.rowGap;
    const blockWidthMm = state.config.blockWidth;
    const blockHeightMm = state.config.blockHeight;

    const fontHostPx = state.config.fontHost;
    const fontAssetPx = state.config.fontAsset;
    const fontSerialPx = state.config.fontSerial;
    const refGapPx = state.config.refGap;
    const blockPaddingPx = state.config.blockPadding;

    const pageMarginPx = pageMarginMm * MM_TO_PX;
    const colGapPx = colGapMm * MM_TO_PX;
    const rowGapPx = rowGapMm * MM_TO_PX;
    const blockWidthPx = blockWidthMm * MM_TO_PX;
    const blockHeightPx = blockHeightMm * MM_TO_PX;

    const refHostHeightPx = fontHostPx * 1.3;
    const itemTotalHeightPx = refHostHeightPx + refGapPx + blockHeightPx;

    const totalRows = Math.ceil(state.dataset.length / gridCols);

    // Calculate Required Sheet Dimensions dynamically based on content
    const totalItemHeightMm = itemTotalHeightPx / MM_TO_PX;
    const requiredHeightMm = (pageMarginMm * 2) + (totalRows * totalItemHeightMm) + Math.max(0, totalRows - 1) * rowGapMm;

    let sheetWidthMm = 210;
    let baseHeightMm = 297;

    if (state.config.pageSize === 'Letter') {
      sheetWidthMm = 215.9;
      baseHeightMm = 279.4;
    } else if (state.config.pageSize === 'Continuous') {
      sheetWidthMm = (pageMarginMm * 2) + (gridCols * blockWidthMm) + Math.max(0, gridCols - 1) * colGapMm;
      baseHeightMm = 0;
    }

    // Dynamic height ensures ALL stickers fit inside SVG without bottom truncation
    const sheetHeightMm = Math.max(baseHeightMm, Math.ceil(requiredHeightMm));

    const sheetWidthPx = sheetWidthMm * MM_TO_PX;
    const sheetHeightPx = sheetHeightMm * MM_TO_PX;

    // Border attribute helper
    let borderAttr = 'stroke="#0f172a" stroke-width="1.5" rx="6" ry="6"';
    if (state.config.borderStyle === 'solid-thin') {
      borderAttr = 'stroke="#0f172a" stroke-width="1"';
    } else if (state.config.borderStyle === 'solid-thick') {
      borderAttr = 'stroke="#0f172a" stroke-width="2"';
    } else if (state.config.borderStyle === 'dashed') {
      borderAttr = 'stroke="#0f172a" stroke-width="1.5" stroke-dasharray="4 3"';
    } else if (state.config.borderStyle === 'none') {
      borderAttr = 'stroke="none"';
    }

    // Text Anchor calculation
    const textAlign = state.config.textAlign;
    let textAnchor = 'middle';
    if (textAlign === 'left') textAnchor = 'start';
    if (textAlign === 'right') textAnchor = 'end';

    let itemsSvg = '';

    state.dataset.forEach((row, index) => {
      const col = index % gridCols;
      const r = Math.floor(index / gridCols);

      const itemX = pageMarginPx + col * (blockWidthPx + colGapPx);
      const itemY = pageMarginPx + r * (itemTotalHeightPx + rowGapPx);

      const hostVal = state.mappings.host ? (row[state.mappings.host] || '') : '';
      const assetVal = state.mappings.asset ? (row[state.mappings.asset] || '') : '';
      const serialVal = state.mappings.serial ? (row[state.mappings.serial] || '') : '';

      const pHost = state.config.prefixHost ? (state.config.prefixHost.endsWith(' ') ? state.config.prefixHost : state.config.prefixHost + ' ') : '';
      const pSerial = state.config.prefixSerial ? (state.config.prefixSerial.endsWith(' ') ? state.config.prefixSerial : state.config.prefixSerial + ' ') : '';

      const hostText = pHost + hostVal;
      const serialText = pSerial + serialVal;

      // 1. Ref Host Text (above block)
      let refTextX = itemX + blockWidthPx / 2;
      if (textAlign === 'left') refTextX = itemX;
      if (textAlign === 'right') refTextX = itemX + blockWidthPx;

      const refHostY = itemY + fontHostPx;

      let itemMarkup = `
    <!-- Sticker Item ${index + 1} -->
    <g class="sticker-item" id="sticker-item-${index + 1}">
      <text x="${refTextX.toFixed(2)}" y="${refHostY.toFixed(2)}" text-anchor="${textAnchor}" font-size="${fontHostPx}px" class="sticker-ref-text">${escapeXML(hostText)}</text>`;

      // 2. Sticker Block Rect
      const blockY = itemY + refHostHeightPx + refGapPx;

      itemMarkup += `
      <rect x="${itemX.toFixed(2)}" y="${blockY.toFixed(2)}" width="${blockWidthPx.toFixed(2)}" height="${blockHeightPx.toFixed(2)}" fill="#ffffff" ${borderAttr} />`;

      // 3. Cut Marks if enabled
      if (state.showCutMarks) {
        const markLen = 6;
        const x1 = itemX;
        const x2 = itemX + blockWidthPx;
        const y1 = blockY;
        const y2 = blockY + blockHeightPx;

        itemMarkup += `
      <!-- Cut Marks -->
      <g class="cut-marks" stroke="#94a3b8" stroke-width="0.75" stroke-dasharray="2 2">
        <line x1="${(x1 - markLen).toFixed(2)}" y1="${y1.toFixed(2)}" x2="${x1.toFixed(2)}" y2="${y1.toFixed(2)}" />
        <line x1="${x1.toFixed(2)}" y1="${(y1 - markLen).toFixed(2)}" x2="${x1.toFixed(2)}" y2="${y1.toFixed(2)}" />
        <line x1="${x2.toFixed(2)}" y1="${y1.toFixed(2)}" x2="${(x2 + markLen).toFixed(2)}" y2="${y1.toFixed(2)}" />
        <line x1="${x2.toFixed(2)}" y1="${(y1 - markLen).toFixed(2)}" x2="${x2.toFixed(2)}" y2="${y1.toFixed(2)}" />
        <line x1="${(x1 - markLen).toFixed(2)}" y1="${y2.toFixed(2)}" x2="${x1.toFixed(2)}" y2="${y2.toFixed(2)}" />
        <line x1="${x1.toFixed(2)}" y1="${y2.toFixed(2)}" x2="${x1.toFixed(2)}" y2="${(y2 + markLen).toFixed(2)}" />
        <line x1="${(x2 + markLen).toFixed(2)}" y1="${y2.toFixed(2)}" x2="${x2.toFixed(2)}" y2="${y2.toFixed(2)}" />
        <line x1="${x2.toFixed(2)}" y1="${y2.toFixed(2)}" x2="${x2.toFixed(2)}" y2="${(y2 + markLen).toFixed(2)}" />
      </g>`;
      }

      // 4. Content inside block
      let contentX = itemX + blockWidthPx / 2;
      if (textAlign === 'left') contentX = itemX + blockPaddingPx + 6;
      if (textAlign === 'right') contentX = itemX + blockWidthPx - blockPaddingPx - 6;

      const hasBarcode = state.config.barcodeToggle !== 'none' && window.JsBarcode;
      const barcodeVal = (state.config.barcodeToggle === 'asset') ? assetVal : (state.config.barcodeToggle === 'serial' ? serialVal : '');

      let assetY = blockY + blockPaddingPx + fontAssetPx + 2;
      let serialY = assetY + fontSerialPx + 4;

      if (hasBarcode && barcodeVal) {
        assetY = blockY + blockPaddingPx + fontAssetPx + 1;
        serialY = assetY + fontSerialPx + 2;
      }

      itemMarkup += `
      <text x="${contentX.toFixed(2)}" y="${assetY.toFixed(2)}" text-anchor="${textAnchor}" font-size="${fontAssetPx}px" class="sticker-asset-text">${escapeXML(assetVal)}</text>
      <text x="${contentX.toFixed(2)}" y="${serialY.toFixed(2)}" text-anchor="${textAnchor}" font-size="${fontSerialPx}px" class="sticker-serial-text">${escapeXML(serialText)}</text>`;

      // 5. Barcode SVG if enabled
      if (hasBarcode && barcodeVal) {
        try {
          const tempSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          JsBarcode(tempSvg, barcodeVal, {
            format: 'CODE128',
            displayValue: false,
            height: 20,
            margin: 0
          });

          const rawBcWidth = parseFloat(tempSvg.getAttribute('width')) || 100;
          const rawBcHeight = parseFloat(tempSvg.getAttribute('height')) || 20;

          const maxBcWidth = blockWidthPx * 0.85;
          const maxBcHeight = Math.max(10, blockHeightPx - (serialY - blockY) - blockPaddingPx - 4);
          
          const scaleX = Math.min(1.2, maxBcWidth / rawBcWidth);
          const scaleY = Math.min(1, maxBcHeight / rawBcHeight);
          const scale = Math.min(scaleX, scaleY);

          const bcScaledWidth = rawBcWidth * scale;
          const bcScaledHeight = rawBcHeight * scale;

          let bcX = itemX + (blockWidthPx - bcScaledWidth) / 2;
          if (textAlign === 'left') bcX = itemX + blockPaddingPx + 4;
          if (textAlign === 'right') bcX = itemX + blockWidthPx - bcScaledWidth - blockPaddingPx - 4;

          const bcY = blockY + blockHeightPx - bcScaledHeight - blockPaddingPx - 2;

          const bcGroupContent = tempSvg.innerHTML;

          itemMarkup += `
      <!-- Barcode Vector -->
      <g transform="translate(${bcX.toFixed(2)}, ${bcY.toFixed(2)}) scale(${scale.toFixed(4)})" class="sticker-barcode">
        ${bcGroupContent}
      </g>`;
        } catch (err) {
          console.warn('SVG Barcode render warning:', err);
        }
      }

      itemMarkup += `
    </g>`;

      itemsSvg += itemMarkup;
    });

    // Assemble Full Sheet SVG Document
    return `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${sheetWidthMm}mm" height="${sheetHeightMm}mm" viewBox="0 0 ${sheetWidthPx.toFixed(2)} ${sheetHeightPx.toFixed(2)}">
  <defs>
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&amp;family=JetBrains+Mono:wght@500;600;700&amp;display=swap');
      .sticker-ref-text {
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
        font-weight: 500;
        fill: #475569;
      }
      .sticker-asset-text {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 700;
        fill: #0f172a;
      }
      .sticker-serial-text {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 600;
        fill: #0f172a;
      }
    </style>
  </defs>
  <rect width="100%" height="100%" fill="#ffffff" />
  ${itemsSvg}
</svg>`;
  }

  // Standalone Single Sticker SVG
  function generateSingleStickerSVG(index) {
    if (!state.dataset || !state.dataset[index]) return '';

    const row = state.dataset[index];
    const MM_TO_PX = 3.7795275591;

    const blockWidthMm = state.config.blockWidth;
    const blockHeightMm = state.config.blockHeight;

    const fontHostPx = state.config.fontHost;
    const fontAssetPx = state.config.fontAsset;
    const fontSerialPx = state.config.fontSerial;
    const refGapPx = state.config.refGap;
    const blockPaddingPx = state.config.blockPadding;

    const blockWidthPx = blockWidthMm * MM_TO_PX;
    const blockHeightPx = blockHeightMm * MM_TO_PX;

    const refHostHeightPx = fontHostPx * 1.3;
    const itemTotalHeightPx = refHostHeightPx + refGapPx + blockHeightPx;

    const paddingPx = 8;
    const svgWidthPx = blockWidthPx + (paddingPx * 2);
    const svgHeightPx = itemTotalHeightPx + (paddingPx * 2);

    const svgWidthMm = (svgWidthPx / MM_TO_PX).toFixed(2);
    const svgHeightMm = (svgHeightPx / MM_TO_PX).toFixed(2);

    const itemX = paddingPx;
    const itemY = paddingPx;

    const hostVal = state.mappings.host ? (row[state.mappings.host] || '') : '';
    const assetVal = state.mappings.asset ? (row[state.mappings.asset] || '') : '';
    const serialVal = state.mappings.serial ? (row[state.mappings.serial] || '') : '';

    const pHost = state.config.prefixHost ? (state.config.prefixHost.endsWith(' ') ? state.config.prefixHost : state.config.prefixHost + ' ') : '';
    const pSerial = state.config.prefixSerial ? (state.config.prefixSerial.endsWith(' ') ? state.config.prefixSerial : state.config.prefixSerial + ' ') : '';

    const hostText = pHost + hostVal;
    const serialText = pSerial + serialVal;

    let borderAttr = 'stroke="#0f172a" stroke-width="1.5" rx="6" ry="6"';
    if (state.config.borderStyle === 'solid-thin') {
      borderAttr = 'stroke="#0f172a" stroke-width="1"';
    } else if (state.config.borderStyle === 'solid-thick') {
      borderAttr = 'stroke="#0f172a" stroke-width="2"';
    } else if (state.config.borderStyle === 'dashed') {
      borderAttr = 'stroke="#0f172a" stroke-width="1.5" stroke-dasharray="4 3"';
    } else if (state.config.borderStyle === 'none') {
      borderAttr = 'stroke="none"';
    }

    const textAlign = state.config.textAlign;
    let textAnchor = 'middle';
    if (textAlign === 'left') textAnchor = 'start';
    if (textAlign === 'right') textAnchor = 'end';

    let refTextX = itemX + blockWidthPx / 2;
    if (textAlign === 'left') refTextX = itemX;
    if (textAlign === 'right') refTextX = itemX + blockWidthPx;

    const refHostY = itemY + fontHostPx;
    const blockY = itemY + refHostHeightPx + refGapPx;

    let contentX = itemX + blockWidthPx / 2;
    if (textAlign === 'left') contentX = itemX + blockPaddingPx + 6;
    if (textAlign === 'right') contentX = itemX + blockWidthPx - blockPaddingPx - 6;

    const hasBarcode = state.config.barcodeToggle !== 'none' && window.JsBarcode;
    const barcodeVal = (state.config.barcodeToggle === 'asset') ? assetVal : (state.config.barcodeToggle === 'serial' ? serialVal : '');

    let assetY = blockY + blockPaddingPx + fontAssetPx + 2;
    let serialY = assetY + fontSerialPx + 4;

    if (hasBarcode && barcodeVal) {
      assetY = blockY + blockPaddingPx + fontAssetPx + 1;
      serialY = assetY + fontSerialPx + 2;
    }

    let barcodeMarkup = '';
    if (hasBarcode && barcodeVal) {
      try {
        const tempSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        JsBarcode(tempSvg, barcodeVal, {
          format: 'CODE128',
          displayValue: false,
          height: 20,
          margin: 0
        });

        const rawBcWidth = parseFloat(tempSvg.getAttribute('width')) || 100;
        const rawBcHeight = parseFloat(tempSvg.getAttribute('height')) || 20;

        const maxBcWidth = blockWidthPx * 0.85;
        const maxBcHeight = Math.max(10, blockHeightPx - (serialY - blockY) - blockPaddingPx - 4);
        
        const scaleX = Math.min(1.2, maxBcWidth / rawBcWidth);
        const scaleY = Math.min(1, maxBcHeight / rawBcHeight);
        const scale = Math.min(scaleX, scaleY);

        const bcScaledWidth = rawBcWidth * scale;
        const bcScaledHeight = rawBcHeight * scale;

        let bcX = itemX + (blockWidthPx - bcScaledWidth) / 2;
        if (textAlign === 'left') bcX = itemX + blockPaddingPx + 4;
        if (textAlign === 'right') bcX = itemX + blockWidthPx - bcScaledWidth - blockPaddingPx - 4;

        const bcY = blockY + blockHeightPx - bcScaledHeight - blockPaddingPx - 2;

        barcodeMarkup = `
    <!-- Barcode Vector -->
    <g transform="translate(${bcX.toFixed(2)}, ${bcY.toFixed(2)}) scale(${scale.toFixed(4)})" class="sticker-barcode">
      ${tempSvg.innerHTML}
    </g>`;
      } catch (err) {
        console.warn('SVG Single Barcode render warning:', err);
      }
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidthMm}mm" height="${svgHeightMm}mm" viewBox="0 0 ${svgWidthPx.toFixed(2)} ${svgHeightPx.toFixed(2)}">
  <defs>
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&amp;family=JetBrains+Mono:wght@500;600;700&amp;display=swap');
      .sticker-ref-text {
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
        font-weight: 500;
        fill: #475569;
      }
      .sticker-asset-text {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 700;
        fill: #0f172a;
      }
      .sticker-serial-text {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 600;
        fill: #0f172a;
      }
    </style>
  </defs>
  <rect width="100%" height="100%" fill="#ffffff" />
  <g class="sticker-item">
    <text x="${refTextX.toFixed(2)}" y="${refHostY.toFixed(2)}" text-anchor="${textAnchor}" font-size="${fontHostPx}px" class="sticker-ref-text">${escapeXML(hostText)}</text>
    <rect x="${itemX.toFixed(2)}" y="${blockY.toFixed(2)}" width="${blockWidthPx.toFixed(2)}" height="${blockHeightPx.toFixed(2)}" fill="#ffffff" ${borderAttr} />
    <text x="${contentX.toFixed(2)}" y="${assetY.toFixed(2)}" text-anchor="${textAnchor}" font-size="${fontAssetPx}px" class="sticker-asset-text">${escapeXML(assetVal)}</text>
    <text x="${contentX.toFixed(2)}" y="${serialY.toFixed(2)}" text-anchor="${textAnchor}" font-size="${fontSerialPx}px" class="sticker-serial-text">${escapeXML(serialText)}</text>
    ${barcodeMarkup}
  </g>
</svg>`;
  }

  // Download / Save helper for SVG Sheet
  function exportSheetAsSVG() {
    const svgString = generateFullSheetSVG();
    if (!svgString) {
      alert('No stickers available to export.');
      return;
    }
    const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    const filename = `IT_Sticker_Sheet_${Date.now()}.svg`;
    if (window.saveAs) {
      window.saveAs(blob, filename);
    } else {
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  }

  // Export Zip of individual SVG files
  async function exportStickersAsSVGZip() {
    if (!window.JSZip || !window.saveAs) {
      alert('JSZip or FileSaver library missing.');
      return;
    }
    if (!state.dataset || state.dataset.length === 0) {
      alert('No stickers to export.');
      return;
    }

    const zip = new JSZip();
    const folder = zip.folder('Asset_Stickers_SVG');

    state.dataset.forEach((row, i) => {
      const svgContent = generateSingleStickerSVG(i);
      const assetVal = state.mappings.asset ? (row[state.mappings.asset] || `sticker_${i+1}`) : `sticker_${i+1}`;
      const cleanName = assetVal.replace(/[^a-zA-Z0-9_-]/g, '_');
      folder.file(`${cleanName}.svg`, svgContent);
    });

    try {
      const content = await zip.generateAsync({ type: 'blob' });
      saveAs(content, `IT_Asset_Stickers_SVG_${Date.now()}.zip`);
    } catch (err) {
      console.error('ZIP SVG export error:', err);
      alert('Failed to generate ZIP of SVG stickers.');
    }
  }

  // ==========================================
  // Modal & Export Event Bindings
  // ==========================================
  function bindModalEvents() {
    const dataModal = document.getElementById('data-modal');
    const exportModal = document.getElementById('export-modal');

    document.getElementById('btn-toggle-table').addEventListener('click', () => {
      dataModal.classList.add('active');
    });
    document.getElementById('btn-close-modal').addEventListener('click', () => {
      dataModal.classList.remove('active');
    });

    const btnExportSvg = document.getElementById('btn-export-svg');
    if (btnExportSvg) {
      btnExportSvg.addEventListener('click', () => {
        exportSheetAsSVG();
      });
    }

    const btnExportTrigger = document.getElementById('btn-export-modal-trigger');
    if (btnExportTrigger) {
      btnExportTrigger.addEventListener('click', () => {
        exportModal.classList.add('active');
      });
    }

    document.getElementById('btn-close-export-modal').addEventListener('click', () => {
      exportModal.classList.remove('active');
    });

    // Close modals when clicking outside
    [dataModal, exportModal].forEach(modal => {
      if (modal) {
        modal.addEventListener('click', (e) => {
          if (e.target === modal) modal.classList.remove('active');
        });
      }
    });

    // Export Options inside Modal
    const btnExpSheetSvg = document.getElementById('btn-exp-sheet-svg');
    if (btnExpSheetSvg) {
      btnExpSheetSvg.addEventListener('click', () => {
        exportModal.classList.remove('active');
        exportSheetAsSVG();
      });
    }

    const btnExpZipSvg = document.getElementById('btn-exp-zip-svg');
    if (btnExpZipSvg) {
      btnExpZipSvg.addEventListener('click', () => {
        exportModal.classList.remove('active');
        exportStickersAsSVGZip();
      });
    }

    const btnExpSheetPng = document.getElementById('btn-exp-sheet-png');
    if (btnExpSheetPng) {
      btnExpSheetPng.addEventListener('click', () => {
        exportModal.classList.remove('active');
        exportSheetAsPNG();
      });
    }

    const btnExpZipPng = document.getElementById('btn-exp-zip-png');
    if (btnExpZipPng) {
      btnExpZipPng.addEventListener('click', () => {
        exportModal.classList.remove('active');
        exportStickersAsZip();
      });
    }
  }

  // Render Data Table in Modal
  function renderTable() {
    const headRow = document.getElementById('table-head-row');
    const bodyRows = document.getElementById('table-body-rows');

    headRow.innerHTML = '<th>#</th>';
    state.columns.forEach(col => {
      const th = document.createElement('th');
      th.textContent = col;
      headRow.appendChild(th);
    });

    bodyRows.innerHTML = '';
    state.dataset.forEach((row, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${idx + 1}</td>`;
      state.columns.forEach(col => {
        const td = document.createElement('td');
        td.textContent = row[col] || '';
        tr.appendChild(td);
      });
      bodyRows.appendChild(tr);
    });
  }

  // Export Entire Sheet as high-res PNG
  function exportSheetAsPNG() {
    if (!window.html2canvas) {
      alert('html2canvas library failed to load.');
      return;
    }
    const currentScale = state.zoom / 100;
    paperSheet.style.transform = 'none'; // reset zoom temporarily for clean capture

    html2canvas(paperSheet, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff'
    }).then(canvas => {
      paperSheet.style.transform = `scale(${currentScale})`;
      const imageURI = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      link.download = `IT_Sticker_Sheet_${Date.now()}.png`;
      link.href = imageURI;
      link.click();
    }).catch(err => {
      paperSheet.style.transform = `scale(${currentScale})`;
      console.error('PNG Export error:', err);
      alert('Failed to generate sheet PNG.');
    });
  }

  // Export each sticker into a ZIP of PNGs
  async function exportStickersAsZip() {
    if (!window.JSZip || !window.html2canvas || !window.saveAs) {
      alert('Required export libraries (JSZip/html2canvas/FileSaver) are missing.');
      return;
    }

    const zip = new JSZip();
    const folder = zip.folder('Asset_Stickers');
    const stickerItems = document.querySelectorAll('.sticker-item-wrapper');

    if (stickerItems.length === 0) {
      alert('No stickers to export.');
      return;
    }

    // Show quick progress feedback
    const btnExp = document.getElementById('btn-exp-zip-png');
    let originalText = '';
    if (btnExp) {
      originalText = btnExp.innerHTML;
      btnExp.innerHTML = `<i class="ri-loader-4-line ri-spin"></i> Zipping ${stickerItems.length} Stickers...`;
      btnExp.disabled = true;
    }

    try {
      for (let i = 0; i < stickerItems.length; i++) {
        const item = stickerItems[i];
        const rowData = state.dataset[i];
        const assetVal = state.mappings.asset ? (rowData[state.mappings.asset] || `sticker_${i+1}`) : `sticker_${i+1}`;
        const cleanName = assetVal.replace(/[^a-zA-Z0-9_-]/g, '_');

        const canvas = await html2canvas(item, {
          scale: 3,
          backgroundColor: '#ffffff'
        });
        
        const dataUrl = canvas.toDataURL('image/png');
        const base64Data = dataUrl.replace(/^data:image\/png;base64,/, '');
        folder.file(`${cleanName}.png`, base64Data, { base64: true });
      }

      const content = await zip.generateAsync({ type: 'blob' });
      saveAs(content, `IT_Asset_Stickers_${Date.now()}.zip`);
    } catch (err) {
      console.error('ZIP generation error:', err);
      alert('Error generating ZIP file.');
    } finally {
      if (btnExp) {
        btnExp.innerHTML = originalText;
        btnExp.disabled = false;
      }
    }
  }

  // ==========================================
  // Zoom Controls
  // ==========================================
  function bindZoomEvents() {
    const zoomValueEl = document.getElementById('zoom-value');
    
    document.getElementById('btn-zoom-in').addEventListener('click', () => {
      if (state.zoom < 200) {
        state.zoom += 15;
        applyZoom();
      }
    });

    document.getElementById('btn-zoom-out').addEventListener('click', () => {
      if (state.zoom > 40) {
        state.zoom -= 15;
        applyZoom();
      }
    });

    document.getElementById('btn-zoom-reset').addEventListener('click', () => {
      state.zoom = 100;
      applyZoom();
    });

    function applyZoom() {
      zoomValueEl.textContent = `${state.zoom}%`;
      paperSheet.style.transform = `scale(${state.zoom / 100})`;
    }
  }

  // Run app initialization
  init();

});
