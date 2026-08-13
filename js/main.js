(function () {
  'use strict';

  const navbar = document.getElementById('navbar');
  const navToggle = document.getElementById('navToggle');
  const navLinks = document.getElementById('navLinks');
  const sections = document.querySelectorAll('section[id]');
  const navAnchors = document.querySelectorAll('.nav-links a[href^="#"]');

  /* ----- Sticky navbar background ----- */
  function handleScroll() {
    navbar.classList.toggle('scrolled', window.scrollY > 40);
    highlightActiveNav();
  }

  window.addEventListener('scroll', handleScroll, { passive: true });
  handleScroll();

  /* ----- Mobile nav toggle ----- */
  navToggle.addEventListener('click', () => {
    const isOpen = navLinks.classList.toggle('open');
    navToggle.classList.toggle('open', isOpen);
    navToggle.setAttribute('aria-expanded', isOpen);
  });

  navAnchors.forEach((link) => {
    link.addEventListener('click', () => {
      navLinks.classList.remove('open');
      navToggle.classList.remove('open');
      navToggle.setAttribute('aria-expanded', 'false');
    });
  });

  /* ----- Active nav link on scroll ----- */
  function highlightActiveNav() {
    const scrollPos = window.scrollY + navbar.offsetHeight + 100;

    sections.forEach((section) => {
      const top = section.offsetTop;
      const height = section.offsetHeight;
      const id = section.getAttribute('id');

      if (scrollPos >= top && scrollPos < top + height) {
        navAnchors.forEach((link) => {
          link.classList.toggle('active', link.getAttribute('href') === `#${id}`);
        });
      }
    });
  }

  /* ----- Intersection Observer: fade + slide-up ----- */
  const revealObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;

        entry.target.classList.add('visible');
        revealObserver.unobserve(entry.target);
      });
    },
    { threshold: 0.12, rootMargin: '0px 0px -40px 0px' }
  );

  document.querySelectorAll('.reveal').forEach((el, index) => {
    if (el.closest('.project-card')) {
      el.style.transitionDelay = `${index * 0.1}s`;
    }
    revealObserver.observe(el);
  });

  /* ----- Staggered project cards ----- */
  const projectCards = document.querySelectorAll('.project-card');
  const projectObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;

        const cards = entry.target.querySelectorAll('.project-card');
        cards.forEach((card, i) => {
          card.style.transitionDelay = `${i * 0.15}s`;
          card.classList.add('visible');
        });
        projectObserver.unobserve(entry.target);
      });
    },
    { threshold: 0.1 }
  );

  const projectsSection = document.getElementById('projects');
  if (projectsSection) {
    projectCards.forEach((card) => {
      card.classList.add('reveal');
    });
    projectObserver.observe(projectsSection);
  }

  /* ----- Animated skill bars ----- */
  const skillObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;

        const fills = entry.target.querySelectorAll('.skill-fill');
        fills.forEach((fill, i) => {
          const width = fill.getAttribute('data-width');
          fill.style.setProperty('--target-width', `${width}%`);
          setTimeout(() => fill.classList.add('animated'), i * 100);
        });
        skillObserver.unobserve(entry.target);
      });
    },
    { threshold: 0.2 }
  );

  document.querySelectorAll('.skill-category').forEach((cat) => {
    skillObserver.observe(cat);
  });

  /* ----- Hero stagger on load ----- */
  window.addEventListener('load', () => {
    const heroReveals = document.querySelectorAll('.hero .reveal');
    heroReveals.forEach((el, i) => {
      setTimeout(() => el.classList.add('visible'), 150 + i * 120);
    });
  });
})();
