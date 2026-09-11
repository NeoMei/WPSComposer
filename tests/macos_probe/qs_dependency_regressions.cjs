'use strict';

// Characterize the security assumptions of the installed transitive dependency,
// including each consumer's resolution so a stale nested qs cannot hide.
const assert = require('node:assert/strict');
const { createRequire } = require('node:module');
const path = require('node:path');
const test = require('node:test');
const probeRequire = createRequire(path.resolve(__dirname, '../../macos/wps-jsapi-probe/package.json'));
const consumers = ['express', 'body-parser'];

for (const consumer of consumers) {
  const consumerRequire = createRequire(probeRequire.resolve(`${consumer}/package.json`));
  const qs = consumerRequire('qs');

  test(`${consumer}: bracket comma arrays obey configured limits`, () => {
    const options = { comma: true, arrayLimit: 3, throwOnLimitExceeded: true };
    for (const query of ['a=1,2,3,4', 'a[]=1,2,3,4', 'a%5B%5D=1,2,3,4']) {
      assert.throws(() => qs.parse(query, options), RangeError, query);
    }
    assert.deepEqual(qs.parse('a=1,2,3', options), { a: ['1', '2', '3'] });
    // Encoded commas are literal data; splitting occurs before decoding.
    assert.deepEqual(qs.parse('a%5B%5D=1%2C2%2C3%2C4', options), { a: ['1,2,3,4'] });
  });

  test(`${consumer}: parsed constructor.isBuffer is data when serialized`, () => {
    const query = 'x%5Bconstructor%5D%5BisBuffer%5D=y';
    for (const options of [{ plainObjects: true }, { allowPrototypes: true }]) {
      const parsed = qs.parse(query, options);
      assert.equal(parsed.x.constructor.isBuffer, 'y');
      assert.equal(qs.stringify(parsed), query);
    }
    assert.equal(qs.stringify({ name: '文档', tag: ['a', 'b'] }),
      'name=%E6%96%87%E6%A1%A3&tag%5B0%5D=a&tag%5B1%5D=b');
    assert.equal(qs.stringify({ value: Buffer.from('ok') }), 'value=ok');
  });
}

test('Express query and urlencoded requests preserve nested fields, arrays and Unicode', async () => {
  const express = probeRequire('express');
  const http = require('node:http');
  const app = express();
  app.use(express.urlencoded({ extended: true }));
  app.all('/echo', (req, res) => res.json({ query: req.query, body: req.body }));
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  try {
    const result = await new Promise((resolve, reject) => {
      const request = http.request({
        host: '127.0.0.1', port: server.address().port, method: 'POST',
        path: '/echo?document%5Btitle%5D=%E6%96%87%E6%A1%A3&tag[]=a&tag[]=b&literal=1,2',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }, (response) => {
        let body = '';
        response.setEncoding('utf8');
        response.on('data', (chunk) => { body += chunk; });
        response.on('end', () => resolve({ status: response.statusCode, body: JSON.parse(body) }));
      });
      request.on('error', reject);
      request.end('options%5Bformat%5D=docx&pages[]=1&pages[]=2');
    });
    assert.deepEqual(result, { status: 200, body: {
      query: { document: { title: '文档' }, tag: ['a', 'b'], literal: '1,2' },
      body: { options: { format: 'docx' }, pages: ['1', '2'] },
    } });
  } finally {
    await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
  }
});
