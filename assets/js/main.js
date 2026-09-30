/**
 * ==========================================================================
 * 王小明 (Ming Wang) - 個人自我介紹網頁互動腳本
 * Bento Grid Interactivity, 3D Tilt, Theme Switcher & Modals
 * ==========================================================================
 */

// 1. 初始化與全域設定
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initTypewriter();
  initBentoTilt();
  initCopyEmail();
  initProjectModal();
  initResumeDownload();
});

/* ==========================================================================
   2. 深色 / 淺色主題切換 (Dark / Light Mode)
   ========================================================================== */
function initTheme() {
  const themeToggleBtn = document.getElementById('theme-toggle-btn');
  if (!themeToggleBtn) return;

  const sunIcon = document.getElementById('theme-icon-sun');
  const moonIcon = document.getElementById('theme-icon-moon');

  // 讀取既有喜好設定或系統偏好
  const savedTheme = localStorage.getItem('site_theme');
  const systemPrefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  const initialTheme = savedTheme || (systemPrefersDark ? 'dark' : 'dark'); // 預設深色質感

  applyTheme(initialTheme);

  themeToggleBtn.addEventListener('click', () => {
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    applyTheme(newTheme);
    localStorage.setItem('site_theme', newTheme);
    showToast(newTheme === 'dark' ? '🌙 已切換為深邃暗黑模式' : '☀️ 已切換為清新明亮模式');
  });

  function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    if (theme === 'light') {
      if (sunIcon) sunIcon.style.display = 'none';
      if (moonIcon) moonIcon.style.display = 'block';
    } else {
      if (sunIcon) sunIcon.style.display = 'block';
      if (moonIcon) moonIcon.style.display = 'none';
    }
  }
}

/* ==========================================================================
   3. 動態打字機標語效果 (Typewriter Effect)
   ========================================================================== */
function initTypewriter() {
  const typewriterElem = document.getElementById('typewriter');
  if (!typewriterElem) return;

  const phrases = [
    "把好想法，化作流暢絲滑的數位體驗 ✨",
    "專注於現代 Web 全端架構與使用者體驗 🚀",
    "熱愛探索新技術，持續學習並構建好產品 💡",
    "2025 應屆畢業生 • 積極尋找軟體開發職缺 💼"
  ];

  let phraseIndex = 0;
  let charIndex = 0;
  let isDeleting = false;
  const typingSpeed = 75;
  const deletingSpeed = 40;
  const pauseTime = 2200;

  function typeLoop() {
    const currentPhrase = phrases[phraseIndex];

    if (!isDeleting) {
      typewriterElem.textContent = currentPhrase.substring(0, charIndex + 1);
      charIndex++;

      if (charIndex === currentPhrase.length) {
        isDeleting = true;
        setTimeout(typeLoop, pauseTime);
        return;
      }
      setTimeout(typeLoop, typingSpeed);
    } else {
      typewriterElem.textContent = currentPhrase.substring(0, charIndex - 1);
      charIndex--;

      if (charIndex === 0) {
        isDeleting = false;
        phraseIndex = (phraseIndex + 1) % phrases.length;
        setTimeout(typeLoop, 400);
        return;
      }
      setTimeout(typeLoop, deletingSpeed);
    }
  }

  typeLoop();
}

/* ==========================================================================
   4. Bento Grid 卡片 3D 微傾斜與滑鼠微光追蹤 (3D Tilt & Glow)
   ========================================================================== */
function initBentoTilt() {
  // 如果是觸控設備則略過 3D 傾斜以保持最佳捲動體驗
  const isTouchDevice = ('ontouchstart' in window) || (navigator.maxTouchPoints > 0);
  if (isTouchDevice) return;

  const cards = document.querySelectorAll('.bento-card');

  cards.forEach(card => {
    let bounds;

    function onMouseEnter() {
      bounds = card.getBoundingClientRect();
      card.style.transition = 'transform 0.1s ease-out, box-shadow 0.2s ease, border-color 0.2s ease';
    }

    function onMouseMove(e) {
      if (!bounds) bounds = card.getBoundingClientRect();

      const mouseX = e.clientX - bounds.left;
      const mouseY = e.clientY - bounds.top;

      // 設定 CSS 變數以產生卡片內部隨滑鼠移動的聚光效果 (Spotlight Glow)
      card.style.setProperty('--mouse-x', `${mouseX}px`);
      card.style.setProperty('--mouse-y', `${mouseY}px`);

      // 計算卡片中心點與旋轉角度 (-6deg 到 6deg 之間，自然細膩)
      const centerX = bounds.width / 2;
      const centerY = bounds.height / 2;
      const rotateX = -((mouseY - centerY) / centerY) * 5;
      const rotateY = ((mouseX - centerX) / centerX) * 5;

      card.style.transform = `perspective(1000px) rotateX(${rotateX.toFixed(2)}deg) rotateY(${rotateY.toFixed(2)}deg) translateZ(6px)`;
    }

    function onMouseLeave() {
      card.style.transition = 'transform 0.5s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.3s ease, border-color 0.3s ease';
      card.style.transform = 'perspective(1000px) rotateX(0deg) rotateY(0deg) translateZ(0px)';
      card.style.setProperty('--mouse-x', `-100px`);
      card.style.setProperty('--mouse-y', `-100px`);
    }

    card.addEventListener('mouseenter', onMouseEnter);
    card.addEventListener('mousemove', onMouseMove);
    card.addEventListener('mouseleave', onMouseLeave);
  });
}

/* ==========================================================================
   5. 一鍵複製 Email 與 Toast 提示
   ========================================================================== */
function initCopyEmail() {
  const emailBox = document.getElementById('copy-email-box');
  const copyBtn = document.getElementById('copy-email-btn');
  const emailTarget = document.getElementById('email-text');

  if (!emailBox && !copyBtn) return;

  const copyAction = async (e) => {
    e.preventDefault();
    const email = emailTarget ? emailTarget.textContent.trim() : 'ming.wang.dev@gmail.com';

    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(email);
      } else {
        // Fallback for older browsers
        const tempInput = document.createElement('input');
        tempInput.value = email;
        document.body.appendChild(tempInput);
        tempInput.select();
        document.execCommand('copy');
        document.body.removeChild(tempInput);
      }
      showToast(`✨ 信箱已複製：${email}，期待與您聯繫！`);
    } catch (err) {
      showToast(`信箱地址：${email}`);
    }
  };

  if (emailBox) emailBox.addEventListener('click', copyAction);
  if (copyBtn) copyBtn.addEventListener('click', copyAction);
}

/* ==========================================================================
   6. 專案詳情彈窗控制 (Project Detail Modal)
   ========================================================================== */
function initProjectModal() {
  const modal = document.getElementById('project-modal');
  const openButtons = document.querySelectorAll('.js-open-modal');
  const closeBtn = document.getElementById('modal-close');

  if (!modal) return;

  function openModal() {
    modal.classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  function closeModal() {
    modal.classList.remove('active');
    document.body.style.overflow = '';
  }

  openButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      openModal();
    });
  });

  if (closeBtn) {
    closeBtn.addEventListener('click', closeModal);
  }

  // 點擊遮罩外圍關閉
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      closeModal();
    }
  });

  // ESC 鍵關閉
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modal.classList.contains('active')) {
      closeModal();
    }
  });
}

/* ==========================================================================
   7. 履歷下載模擬提示
   ========================================================================== */
function initResumeDownload() {
  const resumeButtons = document.querySelectorAll('.js-download-cv');
  resumeButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      showToast('📄 履歷預覽已就緒，可隨時替換為您的專屬 PDF 檔案！');
    });
  });
}

/* ==========================================================================
   8. 輕量通用 Toast 提示元件
   ========================================================================== */
let toastTimeout;
function showToast(message) {
  let toast = document.getElementById('toast-notification');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'toast-notification';
    toast.className = 'toast-notification';
    document.body.appendChild(toast);
  }

  toast.innerHTML = `
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
      <polyline points="22 4 12 14.01 9 11.01"></polyline>
    </svg>
    <span>${message}</span>
  `;

  toast.classList.add('show');

  if (toastTimeout) clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => {
    toast.classList.remove('show');
  }, 3200);
}
