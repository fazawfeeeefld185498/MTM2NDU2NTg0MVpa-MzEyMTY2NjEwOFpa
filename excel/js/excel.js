'use strict';

const A_CHAR_CODE = 'A'.charCodeAt(0);

const ENTER_KEY_CODE = 13;

const COLUMN_HEADERS = document.getElementById("column-headers");
const EXCEL = document.getElementById("excel");

const FORMULA_TEST_REG = /^[A-Z0-9+\-*/]+$/;
const FORMULA_REG = /[A-Z]+|[+\-*/]+|\d+/g;

let columnSize = 0;
let rowSize = 0;

window.onload = function () {
    init();
};

/**
 * Init
 */
function init() {
    /*for (let i = 0; i < 2; i++) {
     addColumn();
     addRow();
     }*/

    columnSize = 2;
    rowSize = 2;

    let textArea = document.getElementsByTagName("textarea");
    for (let i = 0; i < textArea.length; i++) {
        textArea[i].setAttribute("rows", "1");
        textArea[i].onkeydown = keyHandler;
        textArea[i].onfocus = focusHandler;
        textArea[i].onblur = formulaHandler;
    }
}

/**
 * Adds a new column
 */
function addColumn() {
    let td = document.createElement("td");
    td.setAttribute("class", "column-index");
    td.innerHTML = getNextColumnName(++columnSize);
    COLUMN_HEADERS.appendChild(td);
    let children = EXCEL.children;
    for (let i = 1; i < children.length; i++) {
        let tr = children[i];
        let td = createNewCell();
        tr.appendChild(td);
    }
}

/**
 * Adds a new row
 */
function addRow() {
    let tr = document.createElement("tr");
    let td = document.createElement("td");
    td.setAttribute("class", "row-index");
    td.innerHTML = ++rowSize;
    tr.appendChild(td);
    for (let i = 0; i < columnSize; i++) {
        let td = createNewCell();
        tr.appendChild(td);
    }
    EXCEL.appendChild(tr);
}

/**
 * Key handler
 * @param e Event
 * @returns {boolean}
 */
function keyHandler(e) {
    if (e.keyCode == ENTER_KEY_CODE) {
        let element = e.srcElement;
        element.blur();
        return false;
    }
}

/**
 * Focus handler
 * @param e Event
 */
function focusHandler(e) {
    let element = e.srcElement;
    let formula = element.getAttribute("formula");
    if (formula != null) {
        element.value = formula;
    }
}

/**
 * Formula handler
 * @param e Event
 */
function formulaHandler(e) {
    let element = e.srcElement;
    let value = element.value;
    if (value.startsWith("=")) {
        let formula = value.substr(1, value.length - 1).replace(/\s+/g, "").toUpperCase();
        if (!FORMULA_TEST_REG.test(formula)) {
            alert("The formula you typed contains an error.");
            element.blur();
            element.removeAttribute("formula");
            return;
        }
        element.setAttribute("formula", value);
    } else {
        let formula = element.getAttribute("formula");
        if (formula != null) {
            element.removeAttribute("formula");
        }
    }

    let rows = EXCEL.children;
    for (let i = 1; i < rows.length; i++) {
        let cells = rows[i].children;
        for (let j = 1; j < cells.length; j++) {
            let textArea = cells[j].children[0];
            let formula = textArea.getAttribute("formula");
            if (formula != null) {
                formula = formula.substr(1, formula.length - 1).replace(/\s+/g, "").toUpperCase();
                let array = formula.match(FORMULA_REG);
                let error = 0;
                let k = -1;
                while (k < array.length) {
                    let rawX = array[k = k + 1];
                    if (!isNaN(rawX)) {
                        // Skip operator
                        k++;
                        continue;
                    }
                    let x = getXPosFromColumnName(rawX);
                    let y = parseInt(array[k = k + 1]);
                    let operator = array[k = k + 1];
                    if (operator != null && operator.length != 1) {
                        textArea.value = "#OPERATOR!";
                        error = 1;
                        break;
                    }
                    let value = getValueFromCell(x, y);
                    if (value == null) {
                        // Not found value
                        textArea.value = "#VALUE_N!";
                        error = 1;
                        break;
                    } else if (value.length == 0) {
                        value = 0;
                    } else if (isNaN(value)) {
                        textArea.value = "#VALUE!";
                        error = 1;
                        break;
                    }
                    formula = formula.replace(array[k - 2] + array[k - 1], value);
                }
                if (!error) {
                    textArea.value = eval(formula);
                }
            }
        }
    }
}

/**
 * Gets value from cell
 * @param x Column position
 * @param y Row position
 */
function getValueFromCell(x, y) {
    if (EXCEL.children.length > y) {
        let row = EXCEL.children[y];
        if (row.children.length > x) {
            let column = row.children[x];
            let textArea = column.children[0];
            return textArea.value;
        }
    }
    return null;
}

/**
 * Gets x position form column name
 * @param name Column name
 * @returns {number} X position
 */
function getXPosFromColumnName(name) {
    let ret = 0;
    for (let i = 0; i < name.length; i++) {
        let num = name.charCodeAt(name.length - i - 1) - A_CHAR_CODE + 1;
        ret += num * Math.pow(26, i);
    }
    return ret;
}

/**
 * Returns next column name
 * @param value Column size
 * @returns {string} Next column name
 */
function getNextColumnName(value) {
    let ret = "";
    while (--value >= 0) {
        ret = String.fromCharCode(A_CHAR_CODE + (value % 26)) + ret;
        value /= 26;
    }
    return ret;
}

/**
 * Create a new cell
 * @returns {Element} A new cell's td element
 */
function createNewCell() {
    let td = document.createElement("td");
    let textArea = document.createElement("textarea");
    textArea.setAttribute("rows", "1");
    textArea.onkeydown = keyHandler;
    textArea.onfocus = focusHandler;
    textArea.onblur = formulaHandler;
    td.appendChild(textArea);
    return td;
}
