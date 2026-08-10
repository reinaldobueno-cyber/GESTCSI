import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFile } from 'node:fs/promises';

const html = await readFile(new URL('../index.html', import.meta.url), 'utf8');

function createMapContext() {
  const parseStart = html.indexOf('function normalizeMapText');
  const parseEnd = html.indexOf('function getUF');
  const mapStart = html.indexOf('var CIDADES_COORDS');
  const mapEnd = html.indexOf('function formatMapDate');
  assert.ok(parseStart > 0 && parseEnd > parseStart);
  assert.ok(mapStart > 0 && mapEnd > mapStart);

  const storage = new Map();
  const context = vm.createContext({
    console,
    Promise,
    setTimeout,
    clearTimeout,
    localStorage: {
      getItem(key) { return storage.get(key) || null; },
      setItem(key, value) { storage.set(key, value); }
    }
  });

  vm.runInContext(
    html.slice(parseStart, parseEnd) + '\n' + html.slice(mapStart, mapEnd),
    context
  );
  return context;
}

test('resolves current rural aliases and the Paraguay project deterministically', () => {
  const context = createMapContext();

  const cumaru = context.getCidadeCoordInfo('CUMARU-PA');
  assert.deepEqual(Array.from(cumaru.coords), [-7.81, -50.77]);
  assert.equal(cumaru.resolvedLabel, 'Cumaru do Norte - PA');

  const gogo = context.getCidadeCoordInfo('GOGO DA ONCA-PA');
  assert.deepEqual(Array.from(gogo.coords), [-4.29, -47.55]);
  assert.equal(gogo.resolvedLabel, 'Dom Eliseu - PA');

  const paraguai = context.getCidadeCoordInfo('SALTO DEL GUAIRA - PY');
  assert.deepEqual(Array.from(paraguai.coords), [-24.0633, -54.3076]);
});

test('does not reuse a same-name city from another state', () => {
  const context = createMapContext();
  assert.equal(context.getCidadeCoordInfo('SANTA ISABEL - PA'), null);
  assert.equal(context.getCidadeCoordInfo('GOIANIA'), null);
});

test('builds automatic geocoding requests for the expected country', async () => {
  const context = createMapContext();
  let requestedUrl = '';
  context.fetch = async url => {
    requestedUrl = String(url);
    return {
      ok: true,
      async json() {
        return [{
          lat: '-23.7656',
          lon: '-53.3206',
          address: { country_code: 'br', 'ISO3166-2-lvl4': 'BR-PR' }
        }];
      }
    };
  };

  const found = await context.geocodeCidadeAutomatica_(context.parseCidadeInfo('UMUARAMA - PR'));
  assert.equal(found, true);
  assert.match(requestedUrl, /countrycodes=br/);
  assert.match(decodeURIComponent(requestedUrl), /UMUARAMA - PR, Brasil/);

  context.fetch = async url => {
    requestedUrl = String(url);
    return {
      ok: true,
      async json() {
        return [{
          lat: '-25.2637',
          lon: '-57.5759',
          address: { country_code: 'py' }
        }];
      }
    };
  };

  const pyFound = await context.geocodeCidadeAutomatica_(context.parseCidadeInfo('ASUNCION - PY'));
  assert.equal(pyFound, true);
  assert.match(requestedUrl, /countrycodes=py/);
  assert.match(decodeURIComponent(requestedUrl), /ASUNCION, Paraguai/);
});

test('rejects a geocoder result from the wrong country', async () => {
  const context = createMapContext();
  context.fetch = async () => ({
    ok: true,
    async json() {
      return [{
        lat: '-23.5505',
        lon: '-46.6333',
        address: { country_code: 'br', 'ISO3166-2-lvl4': 'BR-SP' }
      }];
    }
  });

  const found = await context.geocodeCidadeAutomatica_(context.parseCidadeInfo('ASUNCION - PY'));
  assert.equal(found, false);
  assert.equal(context.getCachedMapGeocode_(context.parseCidadeInfo('ASUNCION - PY')), null);
});
