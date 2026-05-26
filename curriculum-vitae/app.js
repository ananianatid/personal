// cursor
const cursor = document.getElementById('cursor');
document.addEventListener('mousemove', e => {
  cursor.style.left = e.clientX + 'px';
  cursor.style.top = e.clientY + 'px';
});
document.querySelectorAll('a, button, .tag, .entry, .hobby-list li').forEach(el => {
  el.addEventListener('mouseenter', () => cursor.classList.add('hover'));
  el.addEventListener('mouseleave', () => cursor.classList.remove('hover'));
});

// theme toggle
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
};

const savedTheme = localStorage.getItem('theme') || 'light';
applyTheme(savedTheme);

toggleBtn.addEventListener('click', () => {
  const current = html.getAttribute('data-theme');
  applyTheme(current === 'light' ? 'dark' : 'light');
});
