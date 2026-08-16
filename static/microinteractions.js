(function () {
    'use strict';

    const root = document.documentElement;
    const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    const bootStartedAt = performance.now();
    let interfaceReleased = false;

    function animateDashboardCounters() {
        if (reducedMotion) return;

        document.querySelectorAll('.stats-grid .metric-card strong').forEach(function (counter, index) {
            const original = counter.textContent.trim();
            const match = original.match(/-?[\d.,]+/);
            if (!match) return;

            const target = Number(match[0].replace(',', '.'));
            if (!Number.isFinite(target)) return;

            const suffix = original.slice(match.index + match[0].length);
            const duration = 880 + (index * 70);
            let startedAt = null;

            counter.setAttribute('aria-label', original);
            counter.textContent = '0' + suffix;

            function tick(timestamp) {
                if (startedAt === null) startedAt = timestamp;
                const progress = Math.min((timestamp - startedAt) / duration, 1);
                const eased = 1 - Math.pow(1 - progress, 3);
                const value = Number.isInteger(target)
                    ? Math.round(target * eased)
                    : (target * eased).toFixed(1);

                counter.textContent = value + suffix;
                if (progress < 1) {
                    window.requestAnimationFrame(tick);
                } else {
                    counter.textContent = original;
                }
            }

            window.requestAnimationFrame(tick);
        });
    }

    function prepareViewportAnimations() {
        if (reducedMotion) return;

        const elements = document.querySelectorAll([
            '.object-card',
            '.progress-card',
            '.board-card',
            '.spill-card',
            '.content-card',
            '.form-card',
            '.detail-summary',
            '.import-guide'
        ].join(','));

        elements.forEach(function (element, index) {
            element.classList.add('ui-observe');
            element.style.setProperty('--ui-delay', Math.min(index % 5, 4) * 40 + 'ms');
        });

        if (!('IntersectionObserver' in window)) {
            elements.forEach(function (element) {
                element.classList.add('ui-visible');
            });
            return;
        }

        const observer = new IntersectionObserver(function (entries) {
            entries.forEach(function (entry) {
                if (!entry.isIntersecting) return;
                entry.target.classList.add('ui-visible');
                observer.unobserve(entry.target);
            });
        }, {
            threshold: 0.08,
            rootMargin: '0px 0px -24px 0px'
        });

        elements.forEach(function (element) {
            observer.observe(element);
        });
    }

    function releaseInterface() {
        if (interfaceReleased) return;
        interfaceReleased = true;

        const minimumSkeletonTime = reducedMotion ? 0 : 120;
        const delay = Math.max(0, minimumSkeletonTime - (performance.now() - bootStartedAt));

        window.setTimeout(function () {
            prepareViewportAnimations();
            root.classList.add('ui-ready');
            window.requestAnimationFrame(function () {
                root.classList.remove('ui-loading');
                animateDashboardCounters();
            });
        }, delay);
    }

    if (document.readyState === 'complete') {
        releaseInterface();
    } else {
        window.addEventListener('load', releaseInterface, { once: true });
        window.setTimeout(releaseInterface, 850);
    }
}());
