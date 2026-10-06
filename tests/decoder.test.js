import test from "node:test";
import assert from "node:assert/strict";

import {
  TEST_DEFINITIONS,
  decodeTestUrl,
} from "../src/decoder.js";

const examples = {
  ito: {
    url: "https://psytests.org/result?v=itoAItqccE",
    scores: [4, 4, 5, 4, 5, 9, 9, 8, 4, 2],
  },
  hol: {
    url: "https://psytests.org/result?v=holO1WhH3E",
    scores: [9, 10, 13, 26, 14, 7],
  },
  ddo: {
    url: "https://psytests.org/result?v=ddoB2Xdr",
    scores: [10, 6, 1, 5, 8],
  },
  can: {
    url: "https://psytests.org/result?v=canC6jT1A",
    scores: [93, 23, 25, 25, 20],
  },
};

test("exposes four test definitions in navigation order", () => {
  assert.deepEqual(Object.keys(TEST_DEFINITIONS), ["ito", "hol", "ddo", "can"]);
});

test("uses a separate result workbook name for every test", () => {
  const fileNames = Object.values(TEST_DEFINITIONS).map(
    (definition) => definition.fileName,
  );

  assert.equal(new Set(fileNames).size, 4);
  assert.deepEqual(fileNames, [
    "ito_results.xlsx",
    "holland_results.xlsx",
    "ddo_results.xlsx",
    "caas_results.xlsx",
  ]);
});

for (const [testId, example] of Object.entries(examples)) {
  test(`decodes the verified ${testId} reference link`, () => {
    const result = decodeTestUrl(example.url, testId);
    const definition = TEST_DEFINITIONS[testId];

    assert.deepEqual(
      definition.metrics.map((metric) => result[metric]),
      example.scores,
    );
  });
}

test("decodes the updated longer RIASEC format and ignores trailing values", () => {
  const result = decodeTestUrl(
    "https://psytests.org/result?v=holO1OZbFyq",
    "hol",
  );

  assert.deepEqual(
    TEST_DEFINITIONS.hol.metrics.map((metric) => result[metric]),
    [9, 10, 13, 26, 14, 7],
  );
});

test("rejects a link for a different selected test", () => {
  assert.throws(
    () => decodeTestUrl(examples.hol.url, "ito"),
    /ссылка относится к тесту «Холланд»/i,
  );
});

test("rejects unsupported versions and malformed links", () => {
  assert.throws(
    () => decodeTestUrl("https://psytests.org/result?v=canX6jT1A", "can"),
    /неподдерживаемая версия/i,
  );
  assert.throws(
    () => decodeTestUrl("https://psytests.org/result", "can"),
    /не найден параметр v/i,
  );
});
