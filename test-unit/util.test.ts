import assert from 'node:assert/strict';

import {describe, test} from '@jest/globals';

// eslint-disable-next-line complexity
function columnNumberToString(columnNumber: number): string {
    const alphabet = 'abcdefghijklmnopqrstuvwxyz';
    const list: Array<string> = [...alphabet.toUpperCase()];

    if (columnNumber < 1) {
        throw new Error('[columnNumberToString] value should be 1 or more');
    }

    if (columnNumber >= 702) {
        throw new Error('[columnNumberToString] value should be less then 702');
    }

    if (columnNumber > 26 && columnNumber % 26 === 0) {
        const firstNumber = columnNumber / 26;

        return list[firstNumber - 2] + 'Z';
    }

    if (columnNumber > 26) {
        const firstNumber = Math.floor(columnNumber / 26);

        return list[firstNumber - 1] + columnNumberToString(columnNumber % 26);
    }

    return list[columnNumber - 1];
}

function columnStringToNumber(columnStringRaw: string): number {
    const alphabet = 'abcdefghijklmnopqrstuvwxyz';
    const alphabetLength: number = alphabet.length;
    const list: Array<string> = [...alphabet.toUpperCase()];
    const columnString = columnStringRaw.trim().toUpperCase();

    return [...columnString].reverse().reduce<number>((sum: number, char: string, index: number) => {
        return sum + (list.indexOf(char) + 1) * Math.pow(alphabetLength, index);
    }, 0);
}

describe('util', () => {
    test('columnNumberToString', () => {
        assert.equal(columnNumberToString(1), 'A');
        assert.throws(() => columnNumberToString(0));
        assert.equal(columnNumberToString(25), 'Y');
        assert.equal(columnNumberToString(26), 'Z');
        assert.equal(columnNumberToString(27), 'AA');
        assert.equal(columnNumberToString(28), 'AB');
        assert.equal(columnNumberToString(52), 'AZ');
        assert.equal(columnNumberToString(53), 'BA');
        assert.equal(columnNumberToString(54), 'BB');
        assert.equal(columnNumberToString(55), 'BC');
        assert.equal(columnNumberToString(56), 'BD');
        assert.equal(columnNumberToString(57), 'BE');
        assert.equal(columnNumberToString(58), 'BF');
        assert.equal(columnNumberToString(59), 'BG');
        assert.equal(columnNumberToString(60), 'BH');
        assert.equal(columnNumberToString(58), 'BF');
        assert.equal(columnNumberToString(71), 'BS');
        assert.equal(columnNumberToString(701), 'ZY');
        assert.throws(() => columnNumberToString(702));
    });

    test('columnStringToNumber', () => {
        assert.equal(columnStringToNumber('A'), 1);
        assert.equal(columnStringToNumber('Y'), 25);
        assert.equal(columnStringToNumber('Z'), 26);
        assert.equal(columnStringToNumber('AA'), 27);
        assert.equal(columnStringToNumber('AB'), 28);
        assert.equal(columnStringToNumber('AZ'), 52);
        assert.equal(columnStringToNumber('BA'), 53);
        assert.equal(columnStringToNumber('BB'), 54);
        assert.equal(columnStringToNumber('BC'), 55);
        assert.equal(columnStringToNumber('BD'), 56);
        assert.equal(columnStringToNumber('BE'), 57);
        assert.equal(columnStringToNumber('BF'), 58);
        assert.equal(columnStringToNumber('BG'), 59);
        assert.equal(columnStringToNumber('BH'), 60);
        assert.equal(columnStringToNumber('BS'), 71);
        assert.equal(columnStringToNumber('ZY'), 701);
        assert.equal(columnStringToNumber('ZZ'), 702);
    });

    test('columnStringToNumber <-> columnNumberToString', () => {
        // eslint-disable-next-line no-loops/no-loops
        for (let index = 1; index <= 701; index += 1) {
            assert.equal(
                columnNumberToString(index),
                columnNumberToString(columnStringToNumber(columnNumberToString(index)))
            );
        }
    });
});
