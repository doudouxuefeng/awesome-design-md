// 滚动揭示动画
const revealTargets = [
  '.hero-title .line',
  '.hero-lede',
  '.hero-actions',
  '.hero-stamp',
  '.section-head',
  '.bean',
  '.story-text',
  '.story-frame',
  '.curve',
  '.subscribe-card',
  '.foot-top',
];

// 给目标元素添加 reveal 类（排除 hero-title 的 line，它们用单独的入场动画）
document.querySelectorAll(revealTargets.join(',')).forEach((el) => {
  el.classList.add('reveal');
});

// Hero 标题行 - 入场时逐行揭示
const heroLines = document.querySelectorAll('.hero-title .line');
heroLines.forEach((line, i) => {
  line.style.transition = `opacity 1s var(--ease) ${0.2 + i * 0.15}s, transform 1s var(--ease) ${0.2 + i * 0.15}s`;
  line.style.opacity = '0';
  line.style.transform = 'translateY(110%)';
});

requestAnimationFrame(() => {
  heroLines.forEach((line) => {
    line.style.opacity = '1';
    line.style.transform = 'translateY(0)';
  });
});

// IntersectionObserver 处理滚动揭示
const io = new IntersectionObserver(
  (entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add('is-visible');
        io.unobserve(entry.target);
      }
    });
  },
  { threshold: 0.15, rootMargin: '0px 0px -80px 0px' }
);

document.querySelectorAll('.reveal').forEach((el) => io.observe(el));

// 订阅套餐选择交互
const planOpts = document.querySelectorAll('.plan-opt');
planOpts.forEach((opt) => {
  opt.addEventListener('click', () => {
    planOpts.forEach((o) => {
      const r = o.querySelector('input');
      if (o !== opt) r.checked = false;
    });
  });
});
