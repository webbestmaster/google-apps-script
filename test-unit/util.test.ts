import assert from 'node:assert/strict';

import {describe, test} from '@jest/globals';

function columnNumberToString(columnNumber: number): string {
    const alphabet = 'abcdefghijklmnopqrstuvwxyz';
    const list: Array<string> = [...alphabet.toUpperCase()];

    if (columnNumber < 1) {
        throw new Error('[columnNumberToString] value should be 1 or more');
    }

    if (columnNumber >= 702) {
        throw new Error('[columnNumberToString] value should be less then 702');
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
        assert.equal(columnStringToNumber('BF'), 58);
        assert.equal(columnStringToNumber('BS'), 71);
        assert.equal(columnStringToNumber('ZY'), 701);
        assert.equal(columnStringToNumber('ZZ'), 702);
    });
});
