const ALPHABET =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";
const ENCODED_ALPHABET =
  "0123456789-_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";

const ITO_METRICS = [
  "Экстраверсия",
  "Спонтанность",
  "Агрессивность",
  "Ригидность",
  "Интроверсия",
  "Сензитивность",
  "Тревожность",
  "Лабильность",
  "Шкала искренности",
  "Шкала аггравации",
];

const ITO_ORDER_LOW_TO_HIGH = [
  "Шкала искренности",
  "Лабильность",
  "Тревожность",
  "Сензитивность",
  "Интроверсия",
  "Ригидность",
  "Агрессивность",
  "Спонтанность",
  "Экстраверсия",
  "Шкала аггравации",
];

const RIASEC_METRICS = [
  "Реалистичный (R)",
  "Исследовательский (I)",
  "Артистичный (A)",
  "Социальный (S)",
  "Предприимчивый (E)",
  "Традиционный (C)",
];

const DDO_METRICS = [
  "Человек (Ч)",
  "Природа (П)",
  "Техника (Т)",
  "Знак. система (З)",
  "Худ. образ (Х)",
];

const CAAS_METRICS = [
  "Карьерная адаптивность",
  "Заинтересованность",
  "Контроль",
  "Любознательность",
  "Уверенность",
];

export const TEST_DEFINITIONS = {
  ito: {
    label: "ИТО",
    title: "Индивидуально-типологический опросник",
    description: "Расшифровка эмоционально-личностных шкал ИТО",
    version: "A",
    payloadLength: 6,
    metrics: ITO_METRICS,
    exampleUrl: "https://psytests.org/result?v=itoAItqccE",
    fileName: "ito_results.xlsx",
    includeLevels: true,
  },
  hol: {
    label: "Холланд",
    title: "Профессиональные типы Холланда",
    description: "Результаты по шести типам профессиональных интересов RIASEC",
    version: "O",
    payloadLength: 6,
    metrics: RIASEC_METRICS,
    exampleUrl: "https://psytests.org/result?v=holO1WhH3E",
    fileName: "holland_results.xlsx",
    includeLevels: false,
  },
  ddo: {
    label: "ДДО",
    title: "Дифференциально-диагностический опросник",
    description: "Профессиональные склонности по типологии Е. А. Климова",
    version: "B",
    payloadLength: 4,
    metrics: DDO_METRICS,
    exampleUrl: "https://psytests.org/result?v=ddoB2Xdr",
    fileName: "ddo_results.xlsx",
    includeLevels: false,
  },
  can: {
    label: "CAAS",
    title: "Шкала карьерной адаптивности",
    description: "Общий показатель и четыре ресурса карьерной адаптивности",
    version: "C",
    payloadLength: 5,
    metrics: CAAS_METRICS,
    exampleUrl: "https://psytests.org/result?v=canC6jT1A",
    fileName: "caas_results.xlsx",
    includeLevels: false,
  },
};

const DDO_TOP = {
  G: "0 1 4 2 10", H: "0 10 6 6 12", I: "1 6 8 11 1", J: "2 2 11 2 3",
  K: "2 12 0 6 5", L: "3 8 2 10 7", M: "4 4 5 1 8", N: "5 0 7 5 10",
  O: "5 9 9 9 12", P: "6 5 12 1 1", Q: "7 2 1 5 3", R: "7 11 3 9 5",
  S: "8 7 6 0 7", T: "9 3 8 4 9", U: "9 12 10 8 11", V: "10 9 0 0 0",
  W: "11 5 2 4 2", X: "12 1 4 8 4", Y: "12 10 6 12 6", Z: "0 6 9 3 7",
  a: "1 2 11 7 9", b: "1 12 0 11 11", c: "2 8 3 3 0", d: "3 4 5 7 2",
  e: "4 0 7 11 4", f: "4 9 10 2 6", g: "5 5 12 6 8", h: "6 2 1 10 10",
  i: "6 11 4 1 12", j: "7 7 6 6 1", k: "8 3 8 10 3", l: "8 12 11 1 5",
  m: "9 9 0 5 6", n: "10 5 2 9 8", o: "11 1 5 0 10", p: "11 10 7 4 12",
  q: "12 6 9 9 1", 1: "1 3 10 3 9", 2: "10 6 1 5 8", 3: "6 8 5 7 7",
};

function positiveModulo(value, modulus) {
  return ((value % modulus) + modulus) % modulus;
}

function charValue(char) {
  const value = ALPHABET.indexOf(char);

  if (value === -1) {
    throw new Error(`Недопустимый символ: ${char}`);
  }

  return value >= 52 ? value - 64 : value;
}

function signedValue(value) {
  let result = 0;

  for (const char of value) {
    result = result * 64 + charValue(char);
  }

  return result;
}

function decodeBase64(value) {
  let result = 0;

  for (const char of value) {
    const digit = ENCODED_ALPHABET.indexOf(char);

    if (digit === -1) {
      throw new Error(`Недопустимый символ: ${char}`);
    }

    result = result * 64 + digit;
  }

  return result;
}

function fromDigits(digits, radix) {
  return digits.reduce((result, digit) => result * radix + digit, 0);
}

function parseDigits(value, radix) {
  return fromDigits(value.split(" ").map(Number), radix);
}

function toDigits(value, radix, count) {
  const digits = [];
  let remaining = value;

  for (let index = 0; index < count; index += 1) {
    digits.push(remaining % radix);
    remaining = Math.floor(remaining / radix);
  }

  return digits.reverse();
}

function toVariableDigits(value, radix) {
  if (value === 0) return [0];

  const digits = [];
  let remaining = value;

  while (remaining > 0) {
    digits.push(remaining % radix);
    remaining = Math.floor(remaining / radix);
  }

  return digits.reverse();
}

function decodeIto(payload) {
  const value = positiveModulo(signedValue(payload) + 3089424140, 10 ** 10);
  const digits = String(value).padStart(10, "0").split("").reverse().map(Number);
  const scores = Object.fromEntries(
    ITO_ORDER_LOW_TO_HIGH.map((metric, index) => [metric, digits[index]]),
  );

  return ITO_METRICS.map((metric) => scores[metric]);
}

function decodeTableTest(payload, reference, table, radix, count) {
  const firstCharacter = payload[0];
  const referenceValue = table[firstCharacter];

  if (!referenceValue) {
    throw new Error(`Первый символ «${firstCharacter}» отсутствует в таблице декодирования`);
  }

  const modulus = radix ** count;
  const offset = positiveModulo(
    parseDigits(referenceValue, radix) - signedValue(reference.slice(1)),
    modulus,
  );
  const value = positiveModulo(offset + signedValue(payload.slice(1)), modulus);

  return toDigits(value, radix, count);
}

function decodeRiasec(payload) {
  const digits = toVariableDigits(decodeBase64(payload), 33);

  if (digits.length < RIASEC_METRICS.length + 1) {
    throw new Error("Код слишком короткий для теста Холланда");
  }

  return digits.slice(1, RIASEC_METRICS.length + 1);
}

function decodeCaas(payload) {
  const value = positiveModulo(23252520 + signedValue("6jT1A") - signedValue(payload), 10 ** 8);
  const text = String(value).padStart(8, "0");
  const subscales = [0, 2, 4, 6].map((index) => Number(text.slice(index, index + 2)));

  return [subscales.reduce((total, score) => total + score, 0), ...subscales];
}

const decoders = {
  ito: decodeIto,
  hol: decodeRiasec,
  ddo: (payload) => decodeTableTest(payload, "2Xdr", DDO_TOP, 13, 5),
  can: decodeCaas,
};

export function getLevel(score) {
  if (score <= 1) return "гипоэмотивность";
  if (score <= 4) return "норма";
  if (score <= 7) return "акцентуация";
  return "избыточность";
}

export function decodeTestUrl(url, selectedTestId) {
  const selectedTest = TEST_DEFINITIONS[selectedTestId];

  if (!selectedTest) {
    throw new Error("Неизвестный тест");
  }

  const match = String(url).match(/[?&]v=([A-Za-z0-9_-]+)/);

  if (!match) {
    throw new Error("Не найден параметр v");
  }

  const code = match[1];
  const prefix = code.slice(0, 3);
  const version = code.slice(3, 4);
  const payload = code.slice(4);
  const linkedTest = TEST_DEFINITIONS[prefix];

  if (!linkedTest) {
    throw new Error(`Неизвестный тип теста «${prefix}»`);
  }

  if (prefix !== selectedTestId) {
    throw new Error(`Ссылка относится к тесту «${linkedTest.label}»`);
  }

  if (version !== selectedTest.version) {
    throw new Error(`Неподдерживаемая версия «${version}»`);
  }

  if (
    selectedTestId !== "hol" &&
    payload.length !== selectedTest.payloadLength
  ) {
    throw new Error("Неожиданная длина кода");
  }

  const scores = decoders[selectedTestId](payload);

  return Object.fromEntries(
    selectedTest.metrics.map((metric, index) => [metric, scores[index]]),
  );
}
