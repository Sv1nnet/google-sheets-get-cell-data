const SHEET_ID = '<SHEET_ID>'

async function getCellData(query) {
  let trimmedQuery = query.toLowerCase().trim();;
  const queryArr = trimmedQuery.split(' ');
  let range = '';

  // if provides just a char and a number
  if (queryArr.length < 3) {
    if (queryArr.length === 2) trimmedQuery = trimmedQuery.split(' ').join('');

    const startsWithCharsAndEndsWithDigits = new RegExp(/^[A-Za-z]{1,}[0-9]{1,}$/);
    const correctlyFormattedPosition = trimmedQuery.match(startsWithCharsAndEndsWithDigits);
    if (correctlyFormattedPosition) range = trimmedQuery.toUpperCase();

    const startsWithDigitsAndEndsWithChars = new RegExp(/^[0-9]{1,}[A-Za-z]{1,}$/);
    const reversedCorrectlyFormattedPosition = trimmedQuery.match(startsWithDigitsAndEndsWithChars);
    if (reversedCorrectlyFormattedPosition) {
      const digitsArr = trimmedQuery.match(/\d/g);
      const lastDigit = digitsArr[digitsArr.length - 1];
      const lastIndexOfDigit = trimmedQuery.lastIndexOf(lastDigit);

      const chars = trimmedQuery.substring(lastIndexOfDigit + 1);
      const digits = digitsArr.join('');
      range = `${chars}${digits}`.toUpperCase();
    };
  }
  if (!range) {
    /* 
      Words that can be a column or a row pointer.
      We need to specify them here to make words with
      mistake, so we can find a row and column even if
      user makes mistake.
    */
    const templatesMistypedWords = getCellData.templatesMistypedWords || {
      row: [
        'строк',
        'строка',
        'строке',
        'строки',
        'строчка',
        'строчке',
        'строчки',
        'row',
      ].map(str => createTemplates(str)),
      column: [
        'столбец',
        'столбца',
        'столбе',
        'столба',
        'столб',
        'column',
      ].map(str => createTemplates(str)),
    }

    if (!getCellData.templatesMistypedWords) getCellData.templatesMistypedWords = templatesMistypedWords;

    const rowInQuery = getQueryData(queryArr, 'row', templatesMistypedWords);
    if (!rowInQuery.isFound) throw new Error('Can\'t find a row in provided query');

    const columnInQuery = getQueryData(queryArr, 'column', templatesMistypedWords);
    if (!columnInQuery.isFound) throw new Error('Can\'t find a column in provided query');

    /* 
    Here is how request looks like if it's not like "A1" or "2b".
    n - number (for column it can be a string aswell), row - string, column - string

    n row | n column 
    n row | column n
    row n | n column 
    row n | column n
    n column | n row
    n column | row n
    column n | n row
    column n | row n
    */

    range = findRowAndColumn(rowInQuery, columnInQuery);
  }

  try {
    const { result } = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range,
    });

    return result.values ? result.values[0][0] : undefined;
  } catch (error) {
    throw new Error(error.result.error.message);
  }


  /* ------------- Helpers ------------- */
  /* 
  Я бы их вынес в отдельный модуль (или хотя бы из функции), но по заданию нужна функция,
  которую можно вставить в консоль браузера и там же запустить.
  */

  function findRowAndColumn(rowInQuery, columnInQuery) {
    if (rowInQuery.index < 2) {
      if (rowInQuery.index === 0) {
        rowInQuery.position = parseInt(queryArr[1], 10);
      } else {
        rowInQuery.position = parseInt(queryArr[0], 10);
      }

      if (columnInQuery.index === 3) {
        if (queryArr[2].match(/\d/)) columnInQuery.position = convertNumberPositionToChar(parseInt(queryArr[2], 10));
        else if (queryArr[2].match(/^[A-Za-z]{1,}$/)) columnInQuery.position = queryArr[2].toUpperCase();
        else throw new Error('Provided column position is not currect');
      } else {
        if (queryArr[3].match(/\d/)) columnInQuery.position = convertNumberPositionToChar(parseInt(queryArr[3], 10));
        else if (queryArr[3].match(/^[A-Za-z]{1,}$/)) columnInQuery.position = queryArr[3].toUpperCase();
        else throw new Error('Provided column position is not currect');
      }
    } else {
      if (rowInQuery.index === 2) {
        rowInQuery.position = parseInt(queryArr[3], 10);
      } else {
        rowInQuery.position = parseInt(queryArr[2], 10);
      }

      if (columnInQuery.index === 0) {
        if (queryArr[1].match(/\d/)) columnInQuery.position = convertNumberPositionToChar(parseInt(queryArr[1], 10));
        else if (queryArr[1].match(/^[A-Za-z]{1,}$/)) columnInQuery.position = queryArr[1].toUpperCase();
        else throw new Error('Provided column position is not currect');
      } else {
        if (queryArr[0].match(/\d/)) columnInQuery.position = convertNumberPositionToChar(parseInt(queryArr[0], 10));
        else if (queryArr[0].match(/^[A-Za-z]{1,}$/)) columnInQuery.position = queryArr[0].toUpperCase();
        else throw new Error('Provided column position is not currect');
      }
    }

    return `${columnInQuery.position}${rowInQuery.position}`;
  }

  // Find a row and a column by templates
  function getQueryData(queryArr, dataType, templatesMistypedWords) {
    let dataIndex = -1;
    const dataIsFound = queryArr.some((word, wordIndex) =>
      templatesMistypedWords[dataType].some(
        templates => templates.some((template) => {
          const index = word.indexOf(template);
          const isFound = index !== -1;

          if (isFound) dataIndex = wordIndex;
          return isFound;
        }),
      )
    );

    return { index: dataIndex, isFound: dataIsFound, position: '' }
  }

  // If user provides column position as a number we should convert it to chars
  function convertNumberPositionToChar(position) {
    const numberOfChars = 26;
    const unicodeShift = 64;
    let pos = position <= 0 ? 1 : position;
    let round = pos / numberOfChars;
    let column = '';
    let prefix = '';
    let isLastChar = false;

    // if column position between A and Z
    if (round <= 1) return String.fromCharCode(pos + unicodeShift);
    if (pos > 18278) throw new Error('Column position is out of table capability');

    pos = pos % 26;
    pos = pos === 0 ? 26 : pos;
    column = String.fromCharCode(pos + unicodeShift);

    // if column position between AA and ZZ
    if (round <= 27) {
      isLastChar = !(round + '').includes('.');
      if (isLastChar) round -= 1;
      prefix = String.fromCharCode(round + unicodeShift);
      return prefix + column;
    }

    // If column position between AAA and ZZZ
    round = (position - numberOfChars) / numberOfChars ** 2;
    isLastChar = !(round + '').includes('.');
    if (isLastChar) round -= 1;

    const firstColumnChar = String.fromCharCode(round + unicodeShift);

    round = parseInt(round);
    const secondColumnChar = convertNumberPositionToChar(position - ((numberOfChars ** 2) * round));
    return firstColumnChar + secondColumnChar;
  }

  /*
    Since a user might make a typo we need to create string that
    could match mistyped word. Here are helpers that create
    these possible mistyped words. Typo is the word where a letter
    on left or on the right from its correct position
  */

  /**
 * 
 * @param {string} str string in which we replace a char
 * @param {string} replacement char to replace
 * @param {index} index index of replacing char
 * @returns {string} string with a replaced char
 */
  function replaceCharByIndex(str, replacement, index) {
    const leftPart = str.substring(index, 0);
    const rightPart = str.substring(index + 1);

    return `${leftPart}${replacement}${rightPart}`;
  }

  /**
   * Create a string with a swapped chars by index
   * @param {string} str string in which we swap chars
   * @param {number} index index of a char to swap; if index === 0 then the function returns string passed to arguments
   * @returns {string} string with swapped chars
   */
  function getTemplateForIndexedChar(str, index) {
    if (index === 0) return str;

    const replacement = str[index];
    const toReplace = str[index - 1];
    let result = replaceCharByIndex(str, toReplace, index);
    return replaceCharByIndex(result, replacement, index - 1);
  }

  /**
   * In case a user makes typo we need words that could match type
   * @param {string} str string that can be mistyped
   * @returns {string[]} array of mistyped string
   */
  function createTemplates(str) {
    const lowerCasedStr = str.toLowerCase();

    return Array.prototype.map.call(lowerCasedStr, (_, index) => getTemplateForIndexedChar(lowerCasedStr, index));
  }
};