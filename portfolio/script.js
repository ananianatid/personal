// cursor
const cursor = document.getElementById('cursor');
document.addEventListener('mousemove', e => {
  cursor.style.left = e.clientX + 'px';
  cursor.style.top = e.clientY + 'px';
});
const updateCursorHover = () => {
  document.querySelectorAll('a, button, .project-card, .stat-card, .skill-group, .hobby-card').forEach(el => {
    if (!el.dataset.hoverBound) {
      el.addEventListener('mouseenter', () => cursor.classList.add('hover'));
      el.addEventListener('mouseleave', () => cursor.classList.remove('hover'));
      el.dataset.hoverBound = "true";
    }
  });
};
updateCursorHover();

// nav scroll
window.addEventListener('scroll', () => {
  document.getElementById('navbar').classList.toggle('scrolled', window.scrollY > 60);
});

// reveal on scroll
const reveals = document.querySelectorAll('.reveal');
const observer = new IntersectionObserver(entries => {
  entries.forEach((e, i) => {
    if (e.isIntersecting) {
      setTimeout(() => e.target.classList.add('visible'), i * 80);
      observer.unobserve(e.target);
    }
  });
}, { threshold: 0.12 });
reveals.forEach(el => observer.observe(el));

// ── THEME TOGGLE ──
const toggleBtn = document.getElementById('theme-toggle');
const html = document.documentElement;
const sunIcon = toggleBtn.querySelector('.sun-icon');
const moonIcon = toggleBtn.querySelector('.moon-icon');

const applyTheme = (theme) => {
  html.setAttribute('data-theme', theme);
  if (theme === 'light') {
    sunIcon.style.display = 'none';
    moonIcon.style.display = '';
  } else {
    sunIcon.style.display = '';
    moonIcon.style.display = 'none';
  }
  localStorage.setItem('theme', theme);
  document.body.style.display = 'none';
  document.body.offsetHeight;
  document.body.style.display = '';
};

const savedTheme = localStorage.getItem('theme') || 'light';
applyTheme(savedTheme);

toggleBtn.addEventListener('click', () => {
  const current = html.getAttribute('data-theme');
  const target = current === 'light' ? 'dark' : 'light';
  applyTheme(target);
});

