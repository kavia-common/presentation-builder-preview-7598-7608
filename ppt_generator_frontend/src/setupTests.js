// jest-dom adds custom jest matchers for asserting on DOM nodes.
// allows you to do things like:
// expect(element).toHaveTextContent(/react/i)
// learn more: https://github.com/testing-library/jest-dom
import '@testing-library/jest-dom';

/**
 * Test-only fetch mock:
 * The app loads schemas via fetch('/assets/...') which, under Jest/JSDOM, resolves to
 * http://localhost/assets/... and causes ECONNREFUSED because no server is running.
 *
 * We serve those assets directly from `public/assets` in the repo during tests.
 */
beforeAll(() => {
  const realFetch = global.fetch;

  global.fetch = async (input, init) => {
    const url = typeof input === 'string' ? input : input?.url;
    if (typeof url === 'string' && url.startsWith('/assets/')) {
      // NOTE: Jest runs with CWD at the package root, so `public/assets` is accessible.
      // eslint-disable-next-line global-require
      const fs = require('fs');
      // eslint-disable-next-line global-require
      const path = require('path');

      const rel = url.replace(/^\/assets\//, '');
      const diskPath = path.join(process.cwd(), 'public', 'assets', rel);

      try {
        const body = fs.readFileSync(diskPath, 'utf8');
        return new Response(body, {
          status: 200,
          headers: new Headers({ 'Content-Type': 'application/json' }),
        });
      } catch (e) {
        return new Response(`Not found: ${url}`, { status: 404 });
      }
    }

    return realFetch(input, init);
  };
});
