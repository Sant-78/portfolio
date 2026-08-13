(function () {
  'use strict';

  const container = document.getElementById('resumeProjects');
  if (!container || typeof PORTFOLIO_PROJECTS === 'undefined') return;

  PORTFOLIO_PROJECTS.forEach(function (project) {
    const entry = document.createElement('article');
    entry.className = 'project-entry';

    const meta =
      project.type === 'internal'
        ? project.company + ' — Internal Tool'
        : project.link
          ? 'GitHub: ' + project.link.replace('https://', '')
          : '';

    const metricsHtml = project.metrics
      .map(function (m) {
        return '<li>' + m + '</li>';
      })
      .join('');

    entry.innerHTML =
      '<h3>' +
      project.title +
      '</h3>' +
      (meta ? '<p class="project-meta">' + meta + '</p>' : '') +
      '<p>' +
      project.description +
      '</p>' +
      '<ul>' +
      metricsHtml +
      '</ul>' +
      (project.workflow ? '<p><em>' + project.workflow + '</em></p>' : '') +
      '<p class="project-tools"><strong>Tools:</strong> ' +
      project.tools.join(', ') +
      '</p>';

    container.appendChild(entry);
  });

  if (new URLSearchParams(window.location.search).get('print') === '1') {
    window.addEventListener('load', function () {
      setTimeout(function () {
        window.print();
      }, 400);
    });
  }

  var printBtn = document.getElementById('printResume');
  if (printBtn) {
    printBtn.addEventListener('click', function () {
      window.print();
    });
  }
})();
