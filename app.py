<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
    <title>Pro EduCreator Studio - Ultimate Master Edition</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>

    <style>
        * { box-sizing: border-box; }
        html, body { 
            margin: 0; padding: 0; width: 100%; height: 100%;
            overflow: hidden; background-color: #0b0b0e; color: white; 
            user-select: none; touch-action: none; cursor: default; 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; 
        }
        
        #presentation-area { 
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; 
            display: flex; justify-content: center; 
            align-items: center; background: #030305; overflow: hidden;
        }

        #slide-image { 
            width: 100%; height: 100%; max-width: 100%; max-height: 100%; 
            object-fit: contain; z-index: 1; border-radius: 4px; 
            box-shadow: 0 10px 30px rgba(0,0,0,0.8); 
        }
        
        #drawing-canvas { 
            position: absolute; top: 0; left: 0; width: 100%; height: 100%; 
            z-index: 10; cursor: default; 
        }
        
        #upload-prompt {
            position: absolute; z-index: 1000; display: flex; flex-direction: column;
            align-items: center; justify-content: center; text-align: center;
            background: #18181b; padding: 30px; border-radius: 20px;
            border: 1px solid rgba(255, 255, 255, 0.15);
            box-shadow: 0 20px 50px rgba(0,0,0,0.9); max-width: 400px; width: 90%;
        }

        .ipad-dock { 
            position: absolute; bottom: 12px; left: 15px; z-index: 100; 
            display: none; flex-direction: row; gap: 5px; background: rgba(20, 20, 25, 0.85); 
            padding: 5px 10px; border-radius: 16px; border: 1px solid rgba(255, 255, 255, 0.12); 
            box-shadow: 0 10px 25px rgba(0,0,0,0.8); backdrop-filter: blur(20px);
            align-items: center; max-width: 85vw; overflow-x: auto;
        }
        
        .dock-btn { 
            width: 28px; height: 28px; border-radius: 50%; cursor: pointer; 
            border: 1px solid transparent; transition: all 0.2s; 
            display: flex; justify-content: center; align-items: center; 
            background: rgba(255, 255, 255, 0.06); color: white; font-size: 11px; flex-shrink: 0;
        }
        .dock-btn:hover { background: rgba(255, 255, 255, 0.2); transform: scale(1.05); }
        .dock-btn.active { border-color: #0a84ff; background: #0a84ff; box-shadow: 0 0 8px rgba(10, 132, 255, 0.5); }
        
        .color-picker-wrapper { position: relative; width: 28px; height: 28px; border-radius: 50%; overflow: hidden; display: flex; align-items: center; justify-content: center; background: rgba(255,255,255,0.06); flex-shrink: 0; }
        #color-picker { opacity: 0; position: absolute; top: 0; left: 0; width: 100%; height: 100%; cursor: pointer; z-index: 2; }
        
        .dock-slider { width: 38px; cursor: pointer; accent-color: #0a84ff; height: 3px; border-radius: 2px; background: #666; flex-shrink: 0; margin: 0 2px; }
        .dock-divider { width: 1px; height: 18px; background: rgba(255,255,255,0.15); margin: 0 2px; flex-shrink: 0; }

        .side-nav-btn {
            position: absolute; top: 50%; transform: translateY(-50%); z-index: 90;
            width: 36px; height: 50px; background: rgba(20, 20, 25, 0.6); border: 1px solid rgba(255,255,255,0.1);
            color: white; display: none; justify-content: center; align-items: center; cursor: pointer;
            backdrop-filter: blur(10px); transition: 0.2s; font-size: 14px;
        }
        .side-nav-btn:hover { background: rgba(10, 132, 255, 0.8); }
        #left-nav { left: 0; border-radius: 0 8px 8px 0; }
        #right-nav { right: 0; border-radius: 8px 0 0 8px; }

        .recorder-ui { 
            position: absolute; top: 12px; right: 12px; z-index: 100; 
            display: none; align-items: center; gap: 6px;
        }
        
        .rec-circle-btn {
            width: 30px; height: 30px; border-radius: 50%; border: none; cursor: pointer;
            display: flex; justify-content: center; align-items: center;
            box-shadow: 0 4px 10px rgba(0,0,0,0.4); transition: 0.3s; color: white; font-size: 11px;
        }
        #btn-start { background: #30d158; }
        #btn-start:hover { background: #28b84d; transform: scale(1.05); }
        #btn-stop { background: #ff453a; display: none; animation: pulse 1.5s infinite; }
        #btn-settings { background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2); }
        #btn-settings:hover { background: rgba(255,255,255,0.25); }

        #rec-settings-menu {
            position: absolute; top: 45px; right: 0; width: 300px; max-height: 85vh; overflow-y: auto;
            background: rgba(20, 20, 25, 0.98); border: 1px solid rgba(255, 255, 255, 0.15);
            border-radius: 12px; padding: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.8);
            backdrop-filter: blur(15px); display: none; z-index: 105; flex-direction: column; gap: 8px;
        }
        .rec-setting-item { display: flex; justify-content: space-between; align-items: center; font-size: 11px; color: #ccc; margin-bottom: 2px;}
        .rec-setting-item select, .rec-setting-item input[type="checkbox"] { cursor: pointer; background: #2c2c35; color: white; border: 1px solid #444; border-radius: 4px; padding: 2px 4px; font-size: 10px; }
        .setting-desc { font-size: 9px; color: #888; margin-top: -4px; margin-bottom: 4px; line-height: 1.1; }

        .timer-badge {
            background: rgba(20, 20, 25, 0.85); padding: 2px 6px; border-radius: 12px;
            font-family: monospace; font-size: 10px; font-weight: bold; border: 1px solid rgba(255,255,255,0.12); display: none; backdrop-filter: blur(10px);
        }

        @keyframes pulse { 0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 69, 58, 0.7); } 70% { transform: scale(1.05); box-shadow: 0 0 0 6px rgba(255, 69, 58, 0); } 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(255, 69, 58, 0); } }

        #loading-screen { display: none; position: absolute; z-index: 2000; width: 100%; height: 100%; background: rgba(0,0,0,0.92); flex-direction: column; justify-content: center; align-items: center; }
        .spinner { border: 4px solid rgba(255,255,255,0.2); border-top: 4px solid #0a84ff; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
</head>
<body>

    <div id="loading-screen">
        <div class="spinner mb-3"></div>
        <h2 id="loading-text" class="text-sm font-medium text-white">फाइल हाई-क्वालिटी में प्रोसेस हो रही है...</h2>
    </div>

    <div id="presentation-area">
        <div id="upload-prompt">
            <i class="fas fa-file-powerpoint text-4xl mb-3 text-[#ff453a]"></i>
            <h2 class="text-lg font-bold mb-2 text-white">क्लास नोट्स या PDF/Images चुनें</h2>
            <button type="button" id="browse-btn" class="bg-[#0a84ff] text-white text-xs px-4 py-2 rounded-xl cursor-pointer hover:bg-blue-600 transition font-semibold shadow-lg mt-4">
                <i class="fas fa-folder-open mr-2"></i> फाइल चुनें (Browse Files)
            </button>
            <input type="file" id="slide-upload" accept=".pdf, .ppt, .pptx, image/*" multiple style="display: none;">
        </div>

        <img id="slide-image" src="" style="display: none;">
        <canvas id="drawing-canvas"></canvas>

        <button class="side-nav-btn" id="left-nav"><i class="fas fa-chevron-left"></i></button>
        <button class="side-nav-btn" id="right-nav"><i class="fas fa-chevron-right"></i></button>
    </div>

    <div class="recorder-ui" id="recorder-ui">
        <span class="timer-badge" id="timer-text">00:00</span>
        <button class="rec-circle-btn" id="btn-start" title="Start Recording"><i class="fas fa-video"></i></button>
        <button class="rec-circle-btn" id="btn-stop" title="Stop & Save"><i class="fas fa-stop"></i></button>
        <button class="rec-circle-btn" id="btn-settings" title="Recording Settings"><i class="fas fa-cog"></i></button>

        <div id="rec-settings-menu">
            <div class="text-[11px] font-bold text-white border-b border-white/10 pb-1 mb-1">रिकॉर्डर सेटिंग्स</div>
            <div class="rec-setting-item">
                <span>माइक्रोफोन (Mic Audio):</span>
                <input type="checkbox" id="setting-mic" checked>
            </div>
            <div class="rec-setting-item">
                <span>वीडियो क्वालिटी:</span>
                <select id="setting-quality">
                    <option value="high">1080p (FHD)</option>
                    <option value="medium" selected>720p (HD)</option>
                </select>
            </div>
            
            <!-- PURE NATURAL AUDIO ENGINE -->
            <div class="text-[11px] font-bold text-[#0a84ff] border-t border-white/10 pt-2 mt-1"><i class="fas fa-headphones"></i> नेचुरल ऑडियो मोड</div>
            
            <div class="rec-setting-item">
                <span class="font-semibold text-blue-400">Low-Rumble Filter:</span>
                <input type="checkbox" id="setting-anti-crackle" checked>
            </div>
            <div class="setting-desc">केवल माइक हिलाने की गूंज हटाता है, आवाज़ पूरी तरह नेचुरल रहती है।</div>

            <div class="rec-setting-item">
                <span class="font-semibold text-yellow-400">Gentle Soft Tone:</span>
                <input type="checkbox" id="setting-sl40-eq" checked>
            </div>
            <div class="setting-desc">तीखेपन को रोककर आवाज़ को कानों के लिए सहज और हल्की रखता है।</div>

            <div class="rec-setting-item">
                <span class="font-semibold text-green-400">Comfort Volume Leveler:</span>
                <input type="checkbox" id="setting-compressor" checked>
            </div>
            <div class="setting-desc">आवाज़ के वॉल्यूम को आरामदायक और सही स्तर पर लॉक रखता है।</div>

            <!-- ADVANCED VIDEO SETTINGS -->
            <div class="text-[11px] font-bold text-gray-400 border-t border-white/10 pt-2 mt-1">वीडियो एडवांस ऑप्शंस</div>
            <div class="rec-setting-item">
                <span>फ्रेम रेट (FPS):</span>
                <select id="setting-fps">
                    <option value="15">15 FPS (Ultra Low CPU)</option>
                    <option value="30" selected>30 FPS (Standard)</option>
                    <option value="60">60 FPS (Smooth)</option>
                </select>
            </div>
            <div class="rec-setting-item">
                <span>बिटरेट (Bitrate):</span>
                <select id="setting-bitrate">
                    <option value="1.5">1.5 Mbps</option>
                    <option value="3" selected>3.0 Mbps</option>
                    <option value="5">5.0 Mbps</option>
                </select>
            </div>

            <!-- MANUAL PERMISSION BUTTON -->
            <button id="btn-request-permission" class="mt-2 w-full bg-[#0a84ff] text-white text-[10px] py-1.5 rounded font-semibold hover:bg-blue-600 transition">
                <i class="fas fa-microphone mr-1"></i> माइक परमिशन दें
            </button>
        </div>
    </div>

    <div class="ipad-dock" id="toolbar">
        <span id="slide-counter" class="text-white text-[10px] font-bold px-1 tracking-wider">1/1</span>
        <div class="dock-divider"></div>
        <div class="color-picker-wrapper" title="रंग बदलें (Color)">
            <input type="color" id="color-picker" value="#ffff00">
            <button class="dock-btn" id="color-btn" style="background: #ffff00; color: black; border: none;"><i class="fas fa-palette text-[8px]"></i></button>
        </div>
        <button class="dock-btn active" id="glow-pen-btn" title="Glow Pen"><i class="fas fa-magic text-[9px]"></i></button>
        <button class="dock-btn" id="normal-pen-btn" title="Normal Pen"><i class="fas fa-pen text-[9px]"></i></button>
        <button class="dock-btn" id="highlighter-btn" title="Highlighter"><i class="fas fa-highlighter text-[9px]"></i></button>
        <button class="dock-btn" id="eraser-btn" title="Eraser"><i class="fas fa-eraser text-[9px]"></i></button>
        <input type="range" id="pen-size" min="2" max="50" value="6" title="Pen Size" class="dock-slider">
        <div class="dock-divider"></div>
        <button class="dock-btn" id="clear-btn" title="Clear Slide" style="color: #ff453a;"><i class="fas fa-trash-alt text-[9px]"></i></button>
        <button class="dock-btn" id="fullscreen-btn" title="Full Screen"><i class="fas fa-expand text-[9px]"></i></button>
    </div>

    <script>
        // PDF & UI Setup
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
        let slides = []; let currentSlideIndex = 0; let slideInks = {}; let screenDirty = true;
        
        const browseBtn = document.getElementById('browse-btn');
        const uploadInput = document.getElementById('slide-upload');
        const slideImage = document.getElementById('slide-image');
        const uploadPrompt = document.getElementById('upload-prompt');
        const loadingScreen = document.getElementById('loading-screen');
        const loadingText = document.getElementById('loading-text');

        browseBtn.addEventListener('click', () => uploadInput.click());

        uploadInput.addEventListener('change', async (e) => {
            const files = Array.from(e.target.files);
            if (files.length === 0) return;
            uploadPrompt.style.display = 'none'; loadingScreen.style.display = 'flex';
            
            try {
                const firstFile = files[0]; const fileName = firstFile.name.toLowerCase();
                if (firstFile.type === 'application/pdf' || fileName.endsWith('.pdf')) {
                    loadingText.innerText = "PDF को High Quality Slides में बदला जा रहा है...";
                    const fileReader = new FileReader();
                    fileReader.onload = async function() {
                        const typedarray = new Uint8Array(this.result);
                        const loadingTask = pdfjsLib.getDocument({ data: typedarray });
                        const pdf = await loadingTask.promise;
                        slides = [];
                        for (let i = 1; i <= pdf.numPages; i++) {
                            const page = await pdf.getPage(i);
                            const viewport = page.getViewport({ scale: 2.0 }); 
                            const canvasTemp = document.createElement('canvas');
                            canvasTemp.height = viewport.height; canvasTemp.width = viewport.width;
                            const contextTemp = canvasTemp.getContext('2d');
                            await page.render({ canvasContext: contextTemp, viewport: viewport }).promise;
                            slides.push(canvasTemp.toDataURL('image/png', 0.9));
                        }
                        startApp();
                    };
                    fileReader.readAsArrayBuffer(firstFile);
                } else {
                    slides = files.sort((a,b) => a.name.localeCompare(b.name)).map(f => URL.createObjectURL(f));
                    startApp();
                }
            } catch (err) {
                alert("फाइल इम्पोर्ट करने में समस्या आई।");
                loadingScreen.style.display = 'none'; uploadPrompt.style.display = 'flex';
            }
        });

        function startApp() {
            loadingScreen.style.display = 'none'; slideImage.style.display = 'block';
            document.getElementById('toolbar').style.display = 'flex';
            document.getElementById('recorder-ui').style.display = 'flex';
            document.getElementById('left-nav').style.display = 'flex';
            document.getElementById('right-nav').style.display = 'flex';
            loadSlide(0);
        }

        function loadSlide(index) {
            saveCurrentInk(); currentSlideIndex = index; slideImage.src = slides[currentSlideIndex]; 
            document.getElementById('slide-counter').innerText = `${currentSlideIndex + 1}/${slides.length}`; 
            slideImage.onload = () => { resizeCanvas(); restoreInk(); screenDirty = true; }; 
        }

        document.getElementById('left-nav').addEventListener('click', () => { if (currentSlideIndex > 0) loadSlide(currentSlideIndex - 1); });
        document.getElementById('right-nav').addEventListener('click', () => { if (currentSlideIndex < slides.length - 1) loadSlide(currentSlideIndex + 1); });
        window.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowRight' || e.key === 'PageDown') { if (currentSlideIndex < slides.length - 1) loadSlide(currentSlideIndex + 1); }
            else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { if (currentSlideIndex > 0) loadSlide(currentSlideIndex - 1); }
        });

        const fullscreenBtn = document.getElementById('fullscreen-btn');
        fullscreenBtn.addEventListener('click', () => {
            let docEl = document.documentElement; let isFs = document.fullscreenElement || document.webkitFullscreenElement;
            if (!isFs) { if (docEl.requestFullscreen) docEl.requestFullscreen(); fullscreenBtn.innerHTML = '<i class="fas fa-compress text-[9px]"></i>'; } 
            else { if (document.exitFullscreen) document.exitFullscreen(); fullscreenBtn.innerHTML = '<i class="fas fa-expand text-[9px]"></i>'; }
            setTimeout(() => { screenDirty = true; }, 400);
        });

        // Drawing Engine
        const canvas = document.getElementById('drawing-canvas'); const ctx = canvas.getContext('2d');
        let isDrawing = false, points = [], currentMode = 'glow', currentColor = '#ffff00';
        const sizeSlider = document.getElementById('pen-size'); const colorPicker = document.getElementById('color-picker'); const colorBtn = document.getElementById('color-btn');
        const tools = { 'glow': document.getElementById('glow-pen-btn'), 'normal': document.getElementById('normal-pen-btn'), 'highlighter': document.getElementById('highlighter-btn'), 'eraser': document.getElementById('eraser-btn') };

        function resizeCanvas() { canvas.width = window.innerWidth; canvas.height = window.innerHeight; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; restoreInk(); screenDirty = true; }
        window.addEventListener('resize', resizeCanvas);

        function setActiveTool(mode) {
            currentMode = mode; Object.values(tools).forEach(btn => btn.classList.remove('active')); tools[mode].classList.add('active');
            if(mode === 'highlighter') sizeSlider.value = 30; else if(mode === 'eraser') sizeSlider.value = 40; else sizeSlider.value = 6;
        }
        setActiveTool('glow');

        colorPicker.addEventListener('input', (e) => { currentColor = e.target.value; colorBtn.style.background = currentColor; });
        tools['glow'].addEventListener('click', () => setActiveTool('glow')); tools['normal'].addEventListener('click', () => setActiveTool('normal')); 
        tools['highlighter'].addEventListener('click', () => setActiveTool('highlighter')); tools['eraser'].addEventListener('click', () => setActiveTool('eraser'));
        document.getElementById('clear-btn').addEventListener('click', () => { ctx.clearRect(0, 0, canvas.width, canvas.height); saveCurrentInk(); screenDirty = true; });

        canvas.addEventListener('pointerdown', (e) => { isDrawing = true; points = [{x: e.clientX, y: e.clientY}]; ctx.beginPath(); ctx.moveTo(e.clientX, e.clientY); screenDirty = true; });
        canvas.addEventListener('pointermove', (e) => {
            if (!isDrawing) return; points.push({x: e.clientX, y: e.clientY}); let baseSize = parseInt(sizeSlider.value);
            
            if (currentMode === 'eraser') { ctx.globalCompositeOperation = 'destination-out'; ctx.lineWidth = baseSize; ctx.shadowBlur = 0; ctx.strokeStyle = 'rgba(0,0,0,1)'; } 
            else if (currentMode === 'highlighter') { ctx.globalCompositeOperation = 'source-over'; ctx.lineWidth = baseSize; ctx.shadowBlur = 0; ctx.strokeStyle = currentColor + '55'; } 
            else if (currentMode === 'glow') { ctx.globalCompositeOperation = 'source-over'; ctx.lineWidth = baseSize; ctx.shadowBlur = 15; ctx.shadowColor = currentColor; ctx.strokeStyle = currentColor; } 
            else { ctx.globalCompositeOperation = 'source-over'; ctx.lineWidth = baseSize; ctx.shadowBlur = 0; ctx.strokeStyle = currentColor; }

            if (points.length >= 3) {
                const controlPoint = points[points.length - 2];
                const endPoint = { x: (points[points.length - 2].x + points[points.length - 1].x) / 2, y: (points[points.length - 2].y + points[points.length - 1].y) / 2 };
                ctx.quadraticCurveTo(controlPoint.x, controlPoint.y, endPoint.x, endPoint.y); ctx.stroke(); ctx.beginPath(); ctx.moveTo(endPoint.x, endPoint.y); screenDirty = true;
            }
        });
        canvas.addEventListener('pointerup', () => { isDrawing = false; points = []; saveCurrentInk(); screenDirty = true; });

        function saveCurrentInk() { if(slides.length > 0) slideInks[currentSlideIndex] = canvas.toDataURL(); }
        function restoreInk() { ctx.clearRect(0, 0, canvas.width, canvas.height); if (slideInks[currentSlideIndex]) { let img = new Image(); img.src = slideInks[currentSlideIndex]; img.onload = () => { ctx.drawImage(img, 0, 0); screenDirty = true; }; } }

        // Settings Menu Toggle
        const btnSettings = document.getElementById('btn-settings'); const recSettingsMenu = document.getElementById('rec-settings-menu');
        btnSettings.addEventListener('click', (e) => { e.stopPropagation(); recSettingsMenu.style.display = recSettingsMenu.style.display === 'flex' ? 'none' : 'flex'; });
        window.addEventListener('click', () => { recSettingsMenu.style.display = 'none'; }); recSettingsMenu.addEventListener('click', (e) => e.stopPropagation());

        // Manual Permission Request
        const btnRequestPermission = document.getElementById('btn-request-permission');
        if(btnRequestPermission) {
            btnRequestPermission.addEventListener('click', async () => {
                try {
                    const stream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false });
                    stream.getTracks().forEach(track => track.stop());
                    alert("माइक्रोफोन की अनुमति मिल गई है!");
                } catch (err) {
                    alert("माइक्रोफोन की अनुमति नहीं मिल सकी।");
                }
            });
        }

        // ============================================================
        // PURE & GENTLE AUDIO ENGINE (NATURAL & COMFORTABLE VOLUME)
        // ============================================================
        let mediaRecorder; let recordedChunks = []; let timerInterval, seconds = 0; let recordingLoopActive = false;
        
        const btnStart = document.getElementById('btn-start'); const btnStop = document.getElementById('btn-stop'); const timerText = document.getElementById('timer-text');

        btnStart.addEventListener('click', async () => {
            try {
                const useMic = document.getElementById('setting-mic').checked;
                const quality = document.getElementById('setting-quality').value;
                const targetFps = parseInt(document.getElementById('setting-fps').value) || 30;
                const bitrateMbps = parseFloat(document.getElementById('setting-bitrate').value) || 3;
                
                const useAntiCrackle = document.getElementById('setting-anti-crackle').checked;
                const useSl40Eq = document.getElementById('setting-sl40-eq').checked;
                const useCompressor = document.getElementById('setting-compressor').checked;

                const recCanvas = document.createElement('canvas'); const recCtx = recCanvas.getContext('2d', { alpha: false });
                recCanvas.width = quality === 'high' ? 1920 : 1280; recCanvas.height = quality === 'high' ? 1080 : 720;

                let activeSlideImg = new Image();
                if (slides.length > 0 && slides[currentSlideIndex]) activeSlideImg.src = slides[currentSlideIndex];

                recordingLoopActive = true; screenDirty = true;
                
                function drawRecordingFrame() {
                    if (!recordingLoopActive) return;
                    if (screenDirty) {
                        recCtx.fillStyle = '#030305'; recCtx.fillRect(0, 0, recCanvas.width, recCanvas.height);
                        if (slides.length > 0 && slides[currentSlideIndex]) {
                            if (activeSlideImg.src !== slides[currentSlideIndex]) activeSlideImg.src = slides[currentSlideIndex];
                            if (activeSlideImg.complete && activeSlideImg.naturalWidth) {
                                let hRatio = recCanvas.width / activeSlideImg.naturalWidth; let vRatio = recCanvas.height / activeSlideImg.naturalHeight; let ratio = Math.min(hRatio, vRatio);
                                let centerShiftX = (recCanvas.width - activeSlideImg.naturalWidth * ratio) / 2; let centerShiftY = (recCanvas.height - activeSlideImg.naturalHeight * ratio) / 2;
                                recCtx.drawImage(activeSlideImg, 0, 0, activeSlideImg.naturalWidth, activeSlideImg.naturalHeight, centerShiftX, centerShiftY, activeSlideImg.naturalWidth * ratio, activeSlideImg.naturalHeight * ratio);
                            }
                        }
                        if (canvas) recCtx.drawImage(canvas, 0, 0, canvas.width, canvas.height, 0, 0, recCanvas.width, recCanvas.height);
                        screenDirty = false;
                    }
                    requestAnimationFrame(drawRecordingFrame);
                }
                drawRecordingFrame();

                const videoStream = recCanvas.captureStream(targetFps);
                const finalStream = new MediaStream();
                videoStream.getVideoTracks().forEach(track => finalStream.addTrack(track));

                if (useMic) {
                    const micStream = await navigator.mediaDevices.getUserMedia({ 
                        audio: { 
                            noiseSuppression: true, 
                            echoCancellation: true, 
                            autoGainControl: false, // ऑटो-गेन बंद किया ताकि आवाज़ अपने आप तेज़ न हो
                            sampleRate: 48000 
                        }, 
                        video: false 
                    });

                    const audioCtx = new (window.AudioContext || window.webkitAudioContext)({ sampleRate: 48000 });
                    if (audioCtx.state === 'suspended') await audioCtx.resume();
                    
                    const source = audioCtx.createMediaStreamSource(micStream);
                    const destination = audioCtx.createMediaStreamDestination();

                    // 1. LIGHT HIGH PASS (केवल 70Hz से नीचे की अनावश्यक घड़घड़ाहट हटाएगा)
                    const gentleHighPass = audioCtx.createBiquadFilter();
                    gentleHighPass.type = 'highpass';
                    gentleHighPass.frequency.value = 70;

                    // 2. SOFT DE-HARSHENER (हाई-ट्रिबल और तीखेपन को स्मूथ -3.0dB कम करना)
                    const softTrebleCut = audioCtx.createBiquadFilter();
                    softTrebleCut.type = 'highshelf';
                    softTrebleCut.frequency.value = 5200;
                    softTrebleCut.gain.value = -3.0; 

                    // 3. COMFORT VOLUME CONTROL (ओवरऑल आवाज़ को थोड़ा घटाकर कानों के लिए सहज बनाना)
                    const masterGain = audioCtx.createGain();
                    masterGain.gain.value = 0.72; // वॉल्यूम कम करके आरामदायक स्तर पर लाया गया

                    // 4. GENTLE PEAK LIMITER (आवाज़ को फटने या बहुत तेज़ होने से रोकने के लिए)
                    const gentleCompressor = audioCtx.createDynamicsCompressor();
                    gentleCompressor.threshold.value = -18;
                    gentleCompressor.knee.value = 20;
                    gentleCompressor.ratio.value = 1.5;
                    gentleCompressor.attack.value = 0.03;
                    gentleCompressor.release.value = 0.2;

                    // नेचुरल ऑडियो सिग्नल फ़्लो
                    let currentNode = source;
                    
                    if (useAntiCrackle) {
                        currentNode.connect(gentleHighPass);
                        currentNode = gentleHighPass;
                    }

                    if (useSl40Eq) {
                        currentNode.connect(softTrebleCut);
                        currentNode = softTrebleCut;
                    }

                    currentNode.connect(masterGain);
                    currentNode = masterGain;

                    if (useCompressor) {
                        currentNode.connect(gentleCompressor);
                        currentNode = gentleCompressor;
                    }

                    currentNode.connect(destination);
                    destination.stream.getAudioTracks().forEach(track => finalStream.addTrack(track));
                }

                let mimeType = 'video/webm;codecs=vp9,opus'; let fileExtension = 'webm';
                if (MediaRecorder.isTypeSupported('video/mp4;codecs=avc1.42E01E,aac')) { mimeType = 'video/mp4;codecs=avc1.42E01E,aac'; fileExtension = 'mp4'; } 
                else if (MediaRecorder.isTypeSupported('video/mp4')) { mimeType = 'video/mp4'; fileExtension = 'mp4'; } 
                else if (!MediaRecorder.isTypeSupported(mimeType)) { mimeType = 'video/webm'; }

                mediaRecorder = new MediaRecorder(finalStream, { mimeType: mimeType, videoBitsPerSecond: bitrateMbps * 1000000 });
                recordedChunks = [];
                mediaRecorder.ondataavailable = e => { if (e.data.size > 0) recordedChunks.push(e.data); };

                mediaRecorder.onstop = () => {
                    recordingLoopActive = false; clearInterval(timerInterval);
                    timerText.style.display = 'none'; btnStart.style.display = 'flex'; btnStop.style.display = 'none';
                    seconds = 0; timerText.innerText = "00:00";

                    const blob = new Blob(recordedChunks, { type: mimeType });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a'); a.href = url;
                    a.download = `Pro_Class_Recording_${new Date().getTime()}.${fileExtension}`;
                    a.click(); URL.revokeObjectURL(url);
                    finalStream.getTracks().forEach(track => track.stop());
                };

                mediaRecorder.start(1000); 
                btnStart.style.display = 'none'; btnStop.style.display = 'flex'; timerText.style.display = 'block';
                
                timerInterval = setInterval(() => {
                    seconds++;
                    const mins = String(Math.floor(seconds / 60)).padStart(2, '0'); const secs = String(seconds % 60).padStart(2, '0');
                    timerText.innerText = `${mins}:${secs}`;
                }, 1000);

            } catch (err) {
                alert("रिकॉर्डिंग/माइक में एरर: " + err.message);
                recordingLoopActive = false;
            }
        });

        btnStop.addEventListener('click', () => {
            if (mediaRecorder && mediaRecorder.state !== 'inactive') {
                mediaRecorder.stop();
            }
        });
    </script>
</body>
</html>
