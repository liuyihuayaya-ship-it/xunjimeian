/* ============================================================
   寻迹梅庵 — 交互脚本
   ============================================================ */
(function () {
  'use strict';

  /* ==========================================================
     DOM 引用
     ========================================================== */
  const navbar    = document.getElementById('navbar');
  const navLinks  = document.getElementById('navLinks');
  const hamburger = document.getElementById('hamburger');
  const lightbox  = document.getElementById('lightbox');
  const lbImg     = document.getElementById('lightboxImg');
  const lbCaption = document.getElementById('lightboxCaption');
  const lbClose   = document.getElementById('lightboxClose');

  /* ==========================================================
     导航栏滚动阴影
     ========================================================== */
  function onScroll() {
    navbar.classList.toggle('scrolled', window.scrollY > 10);
  }
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll(); // init

  /* ==========================================================
     移动端汉堡菜单
     ========================================================== */
  hamburger.addEventListener('click', function () {
    const isOpen = navLinks.classList.toggle('open');
    hamburger.classList.toggle('open', isOpen);
    hamburger.setAttribute('aria-expanded', isOpen);
  });

  // 点击导航链接关闭菜单
  navLinks.querySelectorAll('a').forEach(function (link) {
    link.addEventListener('click', function () {
      navLinks.classList.remove('open');
      hamburger.classList.remove('open');
      hamburger.setAttribute('aria-expanded', 'false');
    });
  });

  /* ==========================================================
     滚动渐现动画 — 已禁用，内容直接显示
     ========================================================== */

  /* ==========================================================
     导航栏当前 section 高亮 (Intersection Observer)
     ========================================================== */
  var sections = [];
  var navAs = navLinks.querySelectorAll('a');

  // 收集所有锚点对应的 section
  navAs.forEach(function (a) {
    var href = a.getAttribute('href');
    if (href && href.startsWith('#')) {
      var el = document.querySelector(href);
      if (el) sections.push({ link: a, section: el });
    }
  });

  var navObserver = new IntersectionObserver(function (entries) {
    entries.forEach(function (entry) {
      if (entry.isIntersecting) {
        navAs.forEach(function (a) { a.classList.remove('active'); });
        var match = sections.find(function (s) { return s.section === entry.target; });
        if (match) match.link.classList.add('active');
      }
    });
  }, { threshold: 0.3, rootMargin: '-80px 0px -60% 0px' });

  sections.forEach(function (s) { navObserver.observe(s.section); });

  /* ==========================================================
     灯箱 Lightbox
     ========================================================== */
  var galleryItems = document.querySelectorAll('.gallery-item');

  galleryItems.forEach(function (item) {
    item.addEventListener('click', function () {
      var imgDiv  = item.querySelector('.gallery-img');
      var caption = item.getAttribute('data-caption') || '';

      // 克隆图片区域到灯箱
      var clone = imgDiv.cloneNode(true);
      clone.classList.add('lightbox-img');
      lbImg.innerHTML = '';
      lbImg.appendChild(clone);
      lbCaption.textContent = caption;
      lightbox.classList.add('open');
      lightbox.setAttribute('aria-hidden', 'false');

      // 防止 body 滚动
      document.body.style.overflow = 'hidden';
    });
  });

  function closeLightbox() {
    lightbox.classList.remove('open');
    lightbox.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
    lbImg.innerHTML = '';
  }

  lbClose.addEventListener('click', closeLightbox);

  // 点击背景关闭
  lightbox.addEventListener('click', function (e) {
    if (e.target === lightbox) closeLightbox();
  });

  // Escape 键关闭
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && lightbox.classList.contains('open')) {
      closeLightbox();
    }
  });

  /* ==========================================================
     导航跳转 — 即时切换，不平滑滚动
     ========================================================== */
  document.querySelectorAll('a[href^="#"]').forEach(function (anchor) {
    anchor.addEventListener('click', function (e) {
      var targetId = this.getAttribute('href');
      if (targetId === '#') return;
      var target = document.querySelector(targetId);
      if (!target) return;

      e.preventDefault();
      // 直接跳转，无动画
      var top = target.getBoundingClientRect().top + window.pageYOffset - 64;
      window.scrollTo({ top: top, behavior: 'instant' });

      // 更新 URL hash
      if (history.pushState) {
        history.pushState(null, null, targetId);
      }
    });
  });

  /* ==========================================================
     初始化
     ========================================================== */
  // 确保页面加载时立即显示 hero 内容（防止无 JS 时内容不可见）
  console.log(
    '%c 寻迹梅庵 %c 东南大学 ',
    'color:#b8353a;font-size:1.2em;font-weight:bold;',
    'color:#5c5242;'
  );
  console.log(
    '%c嚼得菜根 · 做得大事 %c— 百年梅庵，弦歌不辍',
    'color:#9b6b3d;',
    'color:#8a8070;'
  );

})();
