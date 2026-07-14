// 0. Custom Alert Modal Override
window._nativeAlert = window.alert;
window.alert = function (message) {
    const alertOverlay = document.getElementById('custom-alert');
    const alertMessage = document.getElementById('custom-alert-message');
    if (alertOverlay && alertMessage) {
        alertMessage.textContent = message;
        alertOverlay.classList.add('active');
        document.body.style.overflow = 'hidden';
    } else {
        console.warn('Custom alert elements not found, falling back to browser alert.');
        window._nativeAlert(message);
    }
};

function closeCustomAlert() {
    const alertOverlay = document.getElementById('custom-alert');
    if (alertOverlay) {
        alertOverlay.classList.remove('active');
        const activeModals = document.querySelectorAll('.mobile-menu.active, .modal.active, .login-modal.active, .quiz-modal.active');
        if (activeModals.length === 0) {
            document.body.style.overflow = '';
        }
    }
}

// 1. Theme Toggle (Light/Dark)
const themeToggle = document.getElementById('theme-toggle');
const body = document.body;
const toggleIcon = themeToggle.querySelector('i');

themeToggle.addEventListener('click', () => {
    if (body.classList.contains('light-mode')) {
        body.classList.remove('light-mode');
        body.classList.add('dark-mode');
        toggleIcon.className = 'fa-solid fa-sun';
    } else {
        body.classList.remove('dark-mode');
        body.classList.add('light-mode');
        toggleIcon.className = 'fa-solid fa-moon';
    }
});

// 2. 免費身心靈影音專區與懸浮影片播放器
const videoModal = document.getElementById('video-modal');
const videoPlayer = document.getElementById('preview-video');
const videoSourceMp4 = document.getElementById('video-source-mp4');
const videoSourceWebm = document.getElementById('video-source-webm');
const videoTag = document.getElementById('video-modal-tag');
const videoTitle = document.getElementById('video-modal-title');
const videoDesc = document.getElementById('video-modal-desc');

const videoData = {
    course: {
        tag: "免費影音 · 系統課程",
        title: "一、身心靈系統課程",
        desc: "本系統課程從身心靈三維度（定、清、補）出發，協助您系統化了解自身生命能階與落地修行步驟。",
        poster: "images for website/course_cover.png",
        videoMp4: "https://assets.mixkit.co/videos/preview/mixkit-meditating-woman-in-a-beautiful-forest-43187-large.mp4",
        videoWebm: "https://upload.wikimedia.org/wikipedia/commons/transcoded/b/b1/Yoga_Class_in_Session.webm/Yoga_Class_in_Session.webm.480p.vp9.webm"
    },
    energy: {
        tag: "免費影音 · 能量補充",
        title: "二、補充心靈能量系列",
        desc: "聽覺與視覺的雙重音頻療癒。結合頌缽與脈輪調頻能量，能快速穩定並補充您的磁場防護罩。",
        poster: "images for website/energy_cover.png",
        videoMp4: "https://assets.mixkit.co/videos/preview/mixkit-stars-in-space-background-1611-large.mp4",
        videoWebm: "https://upload.wikimedia.org/wikipedia/commons/transcoded/f/f1/Aurora_Borealis_from_the_Space_Station.webm/Aurora_Borealis_from_the_Space_Station.webm.480p.vp9.webm"
    },
    knowledge: {
        tag: "免費影音 · 趣味知識",
        title: "三、有趣身心靈知識",
        desc: "帶您快速看懂人類圖、占星盤與生命靈數的基礎解讀法。用最趣味的角度解開靈魂藍圖。",
        poster: "images for website/knowledge_cover.png",
        videoMp4: "https://assets.mixkit.co/videos/preview/mixkit-mysterious-glowing-crescent-moon-in-sky-48666-large.mp4",
        videoWebm: "https://upload.wikimedia.org/wikipedia/commons/transcoded/3/38/Night_Sky_Timelapse_at_Doi_Inthanon_National_Park.webm/Night_Sky_Timelapse_at_Doi_Inthanon_National_Park.webm.480p.vp9.webm"
    }
};

function openVideoModal(type) {
    const data = videoData[type];
    if (!data) return;

    videoTag.innerText = data.tag;
    videoTitle.innerText = data.title;
    videoDesc.innerText = data.desc;

    // 設定影片來源與封面圖 (支援 WebM 與 MP4 fallback 雙通道)
    videoPlayer.removeAttribute('src'); // 移除直接的 src，確保使用內部 source 元素
    videoPlayer.poster = data.poster;
    videoSourceMp4.src = data.videoMp4;
    videoSourceWebm.src = data.videoWebm;
    videoPlayer.load();

    videoModal.classList.add('active');
    document.body.style.overflow = 'hidden';

    // 嘗試自動播放
    videoPlayer.play().catch(error => {
        console.log("Auto-play was prevented. Waiting for user interaction.", error);
    });
}

function closeVideoModal() {
    videoPlayer.pause();
    videoModal.classList.remove('active');
    document.body.style.overflow = '';
}

// 點擊彈窗外部關閉影片彈窗
videoModal.addEventListener('click', (e) => {
    if (e.target === videoModal) {
        closeVideoModal();
    }
});

// 3. 服務矩陣 Modal 彈窗
const modal = document.getElementById('matrix-modal');
const modalSecTag = document.getElementById('modal-sec-tag');
const modalSecTitle = document.getElementById('modal-sec-title');
const modalSecDetails = document.getElementById('modal-sec-details');

document.querySelectorAll('.matrix-cell:not(.empty)').forEach(cell => {
    cell.addEventListener('click', () => {
        const section = cell.getAttribute('data-section');
        const service = cell.getAttribute('data-service');
        const details = cell.getAttribute('data-details');

        modalSecTag.innerText = section;
        modalSecTitle.innerText = service;
        modalSecDetails.innerText = details;

        modal.classList.add('active');
        body.style.overflow = 'hidden'; // 鎖定網頁滾動
    });
});

function closeModal() {
    modal.classList.remove('active');
    body.style.overflow = '';
}

// 點擊彈窗外部關閉
modal.addEventListener('click', (e) => {
    if (e.target === modal) {
        closeModal();
    }
});

// 會員彈窗 (Auth Modal) 控制
const authModal = document.getElementById('auth-modal');
const authBtn = document.getElementById('auth-btn');
let isSignUpMode = false;

authBtn.addEventListener('click', () => {
    authModal.classList.add('active');
    body.style.overflow = 'hidden'; // 鎖定網頁滾動
});

function closeAuthModal() {
    authModal.classList.remove('active');
    body.style.overflow = '';
    // 重設為登入模式
    if (isSignUpMode) toggleAuthMode();
}

// 點擊會員彈窗外部關閉
authModal.addEventListener('click', (e) => {
    if (e.target === authModal) {
        closeAuthModal();
    }
});

function toggleAuthMode() {
    const title = document.getElementById('auth-modal-title');
    const submitBtn = document.querySelector('#auth-form button[type="submit"]');
    const switchText = document.getElementById('auth-switch-text');
    const switchBtn = document.getElementById('auth-switch-btn');

    isSignUpMode = !isSignUpMode;
    if (isSignUpMode) {
        title.textContent = '註冊會員';
        submitBtn.textContent = '註冊';
        switchText.textContent = '已經有帳號了？';
        switchBtn.textContent = '立即登入';
    } else {
        title.textContent = '會員登入';
        submitBtn.textContent = '登入';
        switchText.textContent = '還沒有帳號嗎？';
        switchBtn.textContent = '立即註冊';
    }
}

function handleAuthSubmit(e) {
    e.preventDefault();
    const email = document.getElementById('auth-email').value;
    alert(`已成功以 ${email} 進行會員${isSignUpMode ? '註冊' : '登入'}模擬！`);
    closeAuthModal();
}

function handleSocialLogin(platform) {
    alert(`已觸發 ${platform} 快速登入模擬！`);
    closeAuthModal();
}

// 4. 命盤工具輪播卡片
const track = document.getElementById('carousel-track');
const slides = Array.from(track.children);
const prevBtn = document.getElementById('carousel-prev');
const nextBtn = document.getElementById('carousel-next');
const dotsContainer = document.getElementById('carousel-dots');

let slideWidth = slides[0].getBoundingClientRect().width;
let activeSlideIndex = 0;

// 初始化圓點
slides.forEach((_, idx) => {
    const dot = document.createElement('div');
    dot.classList.add('carousel-dot');
    if (idx === 0) dot.classList.add('active');
    dot.addEventListener('click', () => moveToSlide(idx));
    dotsContainer.appendChild(dot);
});
const dots = Array.from(dotsContainer.children);

// 響應式調整寬度
window.addEventListener('resize', () => {
    slideWidth = slides[0].getBoundingClientRect().width;
    moveToSlide(activeSlideIndex);
});

function moveToSlide(index) {
    track.style.transform = `translateX(-${slideWidth * index}px)`;
    dots[activeSlideIndex].classList.remove('active');
    dots[index].classList.add('active');
    activeSlideIndex = index;
}

prevBtn.addEventListener('click', () => {
    let index = activeSlideIndex - 1;
    if (index < 0) index = slides.length - 1;
    moveToSlide(index);
});

nextBtn.addEventListener('click', () => {
    let index = (activeSlideIndex + 1) % slides.length;
    moveToSlide(index);
});

// 觸控滑動支援 (Swipe)
let startX = 0;
let endX = 0;

track.addEventListener('touchstart', (e) => {
    startX = e.touches[0].clientX;
}, { passive: true });

track.addEventListener('touchend', (e) => {
    endX = e.changedTouches[0].clientX;
    const diffX = startX - endX;

    if (Math.abs(diffX) > 50) { // 滑動距離大於 50px 才觸發
        if (diffX > 0) {
            // 往左滑，下一張
            let index = (activeSlideIndex + 1) % slides.length;
            moveToSlide(index);
        } else {
            // 往右滑，上一張
            let index = activeSlideIndex - 1;
            if (index < 0) index = slides.length - 1;
            moveToSlide(index);
        }
    }
}, { passive: true });

// 5. 免費命理測驗
function calculateSoulProfile(event) {
    event.preventDefault();

    const name = document.getElementById('user-name').value;
    const dateStr = document.getElementById('user-date').value;
    const time = document.getElementById('user-time').value;
    const location = document.getElementById('user-location').value;

    if (!name || !dateStr || !time || !location) return;

    document.getElementById('quiz-form-container').style.display = 'none';
    const loading = document.getElementById('quiz-loading');
    loading.style.display = 'flex';

    // 模擬分析載入
    setTimeout(() => {
        loading.style.display = 'none';
        const result = document.getElementById('quiz-result');
        result.style.display = 'block';

        // 根據輸入計算簡單卻看起來專業的交叉結果
        const birthYear = new Date(dateStr).getFullYear();
        const birthMonth = new Date(dateStr).getMonth() + 1;
        const birthDay = new Date(dateStr).getDate();

        // 1. 生命靈數主命數計算 (數字全部相加至個位數)
        const dateDigits = (name.length + birthYear + birthMonth + birthDay).toString().split('');
        let sum = dateDigits.reduce((acc, curr) => acc + parseInt(curr), 0);
        while (sum > 9 && sum !== 11 && sum !== 22 && sum !== 33) {
            sum = sum.toString().split('').reduce((acc, curr) => acc + parseInt(curr), 0);
        }

        // 2. 元素計算 (根據月份與時間)
        const elements = ["烈火系 (火象)", "磐石系 (土象)", "微風系 (風象)", "流水系 (水象)"];
        const hour = parseInt(time.split(':')[0]) || 0;
        const elementIndex = (birthMonth + hour) % 4;
        const element = elements[elementIndex];

        // 3. 人類圖人生角色
        const profiles = ["1/3 探索探究者", "2/4 隱士傳播者", "3/5 烈士異端者", "4/6 傳播觀察者", "5/1 異端探究者", "6/2 典範隱士者"];
        const profile = profiles[(birthDay + name.length) % profiles.length];

        // 4. 紫微主星
        const ziweiStars = ["紫微天府 · 帝后雙尊格", "武曲七殺 · 開疆闢土格", "天機太陰 · 智謀清雅格", "太陽巨門 · 聲名遠播格", "天同天梁 · 因果福庇格", "廉貞貪狼 · 桃花才藝格"];
        const ziwei = ziweiStars[(birthYear + birthMonth) % ziweiStars.length];

        // 5. 撰寫客製化結論
        let conclusion = `親愛的 ${name}，您的出生星盤顯示您屬於【${element}】。這代表您天生具備`;
        if (elementIndex === 0) conclusion += "極佳的直覺與熱情驅動力，對新興事物反應極快。";
        else if (elementIndex === 1) conclusion += "踏實穩定的務實感與承載力，是團隊中不可或缺的基石。";
        else if (elementIndex === 2) conclusion += "敏捷的邏輯思維與資訊整合能力，天生擅長傳播理念。";
        else conclusion += "強大的情感共情力與藝術感知直覺，能輕易同理他人的隱密痛點。";

        conclusion += `結合人類圖 ${profile} 與紫微斗數中的【${ziwei}】格局，您在世俗社會最合適扮演`;

        if (sum === 8 || sum === 11 || sum === 22) {
            conclusion += `「高階管理者與商業變現引領者」的角色。生命靈數 ${sum} 賦予您天生掌控大局的變現能量。然而，您也容易因為過度操勞肉體（身）或執著於金錢指標，導致能量卡關。`;
        } else {
            conclusion += `「智囊顧問、內容創作者與心靈轉化啟盟者」的角色。生命靈數 ${sum} 顯示您此生的核心課題在於追求靈性的圓滿與智慧的開拓。`;
        }

        conclusion += ` 我們建議您當前的「定、清、補」修行步驟為：【定位】先釐清自己的定位與目標，找到自己擅長並且熱愛的事情；【清心】透過鈦缽音療，清除潛意識對未知金錢關卡的恐懼；【補財富】積極補足高級業務與成交心法，將您的靈性天賦與真實世界進行穩固的對齊變現。`;

        // 填入 HTML
        document.getElementById('res-birth-info').innerText = `${name} · 出生西元 ${birthYear}年${birthMonth}月${birthDay}日 ${time} · 出生地${location}`;
        document.getElementById('res-element').innerText = element;
        document.getElementById('res-profile').innerText = profile;
        document.getElementById('res-number').innerText = `主命數 ${sum} · ${sum === 8 || sum === 11 || sum === 22 || sum === 33 ? '高頻變現者' : '靈性覺察者'}`;
        document.getElementById('res-ziwei').innerText = ziwei;
        document.getElementById('res-summary-content').innerText = conclusion;

    }, 2500);
}

function resetQuiz() {
    document.getElementById('quiz-result').style.display = 'none';
    document.getElementById('quiz-form-container').style.display = 'block';
    document.getElementById('soul-quiz-form').reset();
}

// 6. 相片上傳與進度模擬
const uploadZone = document.getElementById('upload-zone');
const progressBox = document.getElementById('upload-progress-box');
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');
const previewBox = document.getElementById('file-preview');
const previewImg = document.getElementById('preview-img');
const previewName = document.getElementById('preview-name');
const previewSize = document.getElementById('preview-size');

let selectedFile = null;

// 拖拽事件
uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadZone.style.borderColor = 'var(--color-primary)';
    uploadZone.style.backgroundColor = 'var(--color-surface)';
});

uploadZone.addEventListener('dragleave', () => {
    uploadZone.style.borderColor = 'var(--color-border)';
    uploadZone.style.backgroundColor = 'var(--color-bg)';
});

uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.style.borderColor = 'var(--color-border)';
    uploadZone.style.backgroundColor = 'var(--color-bg)';
    if (e.dataTransfer.files.length > 0) {
        processUploadedFile(e.dataTransfer.files[0]);
    }
});

function handleFileSelect(event) {
    if (event.target.files.length > 0) {
        processUploadedFile(event.target.files[0]);
    }
}

function processUploadedFile(file) {
    if (!file.type.startsWith('image/')) {
        alert('請上傳有效的相片檔案！');
        return;
    }
    if (file.size > 10 * 1024 * 1024) {
        alert('檔案大小不能超過 10MB！');
        return;
    }

    selectedFile = file;
    uploadZone.style.display = 'none';
    progressBox.style.display = 'block';

    // 模擬上傳與靈性信道對齊進度
    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.floor(Math.random() * 15) + 5;
        if (progress >= 100) {
            progress = 100;
            clearInterval(interval);

            // 顯示預覽
            const reader = new FileReader();
            reader.onload = (e) => {
                previewImg.src = e.target.result;
                previewName.innerText = file.name;
                previewSize.innerText = (file.size / (1024 * 1024)).toFixed(2) + ' MB';

                progressBox.style.display = 'none';
                previewBox.style.display = 'flex';
            };
            reader.readAsDataURL(file);
        }
        progressFill.style.width = progress + '%';
        progressText.innerText = `正在對齊高我信道與相片氣場解構... ${progress}%`;
    }, 150);
}

function clearUploadedFile(event) {
    event.stopPropagation();
    selectedFile = null;
    previewBox.style.display = 'none';
    uploadZone.style.display = 'block';
    document.getElementById('file-input').value = '';
}

// 7. 訂購方案
function purchasePlan(planName) {
    if (!selectedFile) {
        alert('為了讓導師看相接訊解析，請在第一步先上傳您的正面清晰照片！');
        document.getElementById('pricing').scrollIntoView({ behavior: 'smooth' });
        return;
    }

    const email = prompt(`您選擇了【${planName}】方案。\n為了將最終報告書寄送給您，請輸入您的收件 Email 郵箱：`);
    if (email === null) return; // 點擊取消

    if (!email || !email.includes('@')) {
        alert('請輸入正確的 Email 地址！');
        return;
    }

    alert(`【訂購成功！】\n感謝您信任真靈光。\n我們已成功收到您的相片與 Email: ${email}。\n靈性導師將在接訊撰寫完成後，於規定工作天內將報告寄送至您的信箱，請密切注意收信！`);
}

// 8. 滾動顯示動畫
const revealElements = document.querySelectorAll('.stack-card, .editorial-visual, .matrix-container, .carousel-outer, .quiz-card, .upload-container, .pricing-card, .inspiration-box');
const revealOnScroll = () => {
    for (let i = 0; i < revealElements.length; i++) {
        const windowHeight = window.innerHeight;
        const elementTop = revealElements[i].getBoundingClientRect().top;
        const elementVisible = 100;

        if (elementTop < windowHeight - elementVisible) {
            revealElements[i].classList.add('active');
        }
    }
};

window.addEventListener('scroll', revealOnScroll);
revealOnScroll(); // 頁面載入時先執行一次

// 手機版漢堡選單控制
const menuToggle = document.getElementById('menu-toggle');
const mobileMenu = document.getElementById('mobile-menu');
const mobileMenuClose = document.getElementById('mobile-menu-close');
const mobileNavLinks = document.querySelectorAll('.mobile-nav-link');

menuToggle.addEventListener('click', () => {
    mobileMenu.classList.add('active');
    body.style.overflow = 'hidden'; // 防止背景滾動
});

const closeMobileMenu = () => {
    mobileMenu.classList.remove('active');
    body.style.overflow = '';
};

mobileMenuClose.addEventListener('click', closeMobileMenu);

mobileNavLinks.forEach(link => {
    link.addEventListener('click', closeMobileMenu);
});

// 手機版矩陣 Tab 切換與 Modal 啟動控制
function switchMatrixTab(tabName) {
    // 切換按鈕 active 樣式
    const tabBtns = document.querySelectorAll('.matrix-tab-btn');
    tabBtns.forEach(btn => {
        if (btn.getAttribute('onclick').includes(tabName)) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });

    // 切換內容 active 樣式
    const tabContents = document.querySelectorAll('.matrix-tab-content');
    tabContents.forEach(content => {
        if (content.id === `tab-${tabName}`) {
            content.classList.add('active');
        } else {
            content.classList.remove('active');
        }
    });
}

function openMatrixModal(section, service, details) {
    modalSecTag.innerText = section;
    modalSecTitle.innerText = service;
    modalSecDetails.innerText = details;

    modal.classList.add('active');
    body.style.overflow = 'hidden';
}
