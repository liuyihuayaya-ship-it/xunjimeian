/* ============================================================
   寻迹梅庵 — 分页切换脚本
   ============================================================ */
(function () {
  'use strict';

  /* ==========================================================
     DOM 引用
     ========================================================== */
  var navbar    = document.getElementById('navbar');
  var navLinks  = document.getElementById('navLinks');
  var hamburger = document.getElementById('hamburger');
  var lightbox  = document.getElementById('lightbox');
  var lbImg     = document.getElementById('lightboxImg');
  var lbCaption = document.getElementById('lightboxCaption');
  var lbClose   = document.getElementById('lightboxClose');
  var allPages  = document.querySelectorAll('.page');
  var allNavAs  = navLinks.querySelectorAll('a[data-page]');

  /* ==========================================================
     页面切换核心
     ========================================================== */
  function switchPage(pageId) {
    // 隐藏所有页面
    allPages.forEach(function (p) {
      p.classList.remove('active');
    });
    // 显示目标页面
    var target = document.getElementById(pageId);
    if (target) {
      target.classList.add('active');
      target.scrollTop = 0; // 回到顶部
    }
    // 更新导航高亮
    allNavAs.forEach(function (a) {
      a.classList.toggle('active', a.getAttribute('data-page') === pageId);
    });
  }

  // 所有带 data-page 的链接（导航栏 + hero按钮 + 页脚）
  document.querySelectorAll('[data-page]').forEach(function (el) {
    el.addEventListener('click', function (e) {
      e.preventDefault();
      var pageId = this.getAttribute('data-page');
      if (pageId) switchPage(pageId);
      // 移动端关闭菜单
      navLinks.classList.remove('open');
      hamburger.classList.remove('open');
      hamburger.setAttribute('aria-expanded', 'false');
    });
  });

  /* ==========================================================
     导航栏滚动阴影
     ========================================================== */
  function updateNavShadow() {
    var activePage = document.querySelector('.page.active');
    var scrolled = false;
    if (activePage && activePage !== document.getElementById('page-hero')) {
      scrolled = activePage.scrollTop > 10;
    }
    navbar.classList.toggle('scrolled', scrolled);
  }

  // 监听各内容页的滚动
  document.querySelectorAll('.page').forEach(function (page) {
    if (!page.classList.contains('page-hero')) {
      page.addEventListener('scroll', updateNavShadow, { passive: true });
    }
  });

  /* ==========================================================
     移动端汉堡菜单
     ========================================================== */
  hamburger.addEventListener('click', function () {
    var isOpen = navLinks.classList.toggle('open');
    hamburger.classList.toggle('open', isOpen);
    hamburger.setAttribute('aria-expanded', isOpen);
  });

  /* ==========================================================
     灯箱 Lightbox
     ========================================================== */
  var galleryItems = document.querySelectorAll('.gallery-item');

  galleryItems.forEach(function (item) {
    item.addEventListener('click', function () {
      var imgDiv  = item.querySelector('.gallery-img');
      var caption = item.getAttribute('data-caption') || '';

      var clone = imgDiv.cloneNode(true);
      clone.classList.add('lightbox-img');
      lbImg.innerHTML = '';
      lbImg.appendChild(clone);
      lbCaption.textContent = caption;
      lightbox.classList.add('open');
      lightbox.setAttribute('aria-hidden', 'false');
    });
  });

  function closeLightbox() {
    lightbox.classList.remove('open');
    lightbox.setAttribute('aria-hidden', 'true');
    lbImg.innerHTML = '';
  }

  lbClose.addEventListener('click', closeLightbox);

  lightbox.addEventListener('click', function (e) {
    if (e.target === lightbox) closeLightbox();
  });

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && lightbox.classList.contains('open')) {
      closeLightbox();
    }
  });

  /* ==========================================================
     初始化
     ========================================================== */
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
