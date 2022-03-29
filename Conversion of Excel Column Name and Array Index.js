var input = "AC";

console.log(convertExcelColOrIndex(input));

function convertExcelColOrIndex(input) {

    if (!isNaN(input)) {
        return numToSSColumn(input + 1); 
    } else {
        return convertLetterToNumber(input);
    }

    function convertLetterToNumber(str) {
        str = str.toUpperCase();
        let out = 0,
            len = str.length;
        for (pos = 0; pos < len; pos++) {
            out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - pos - 1);
        }
        return out - 1;
    }

    function numToSSColumn(num) {
        let s = '',
            t;

        while (num > 0) {
            t = (num - 1) % 26;
            s = String.fromCharCode(65 + t) + s;
            num = (num - t) / 26 | 0;
        }
        return s || undefined;
    }
}