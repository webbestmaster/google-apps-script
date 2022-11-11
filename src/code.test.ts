import {describe, expect, test} from '@jest/globals';

// import {greeter} from './code';

describe('some test', () => {
    test('just test', async () => {
        expect(1 + 1).toEqual(2);
    });
    test('test greeter', async () => {
        // expect(greeter('user')).toEqual(2);
    });
});
