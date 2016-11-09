# node-xlsx-json

[![Build Status](https://travis-ci.org/totuworld/node-xlsx-json.svg?branch=master)](https://travis-ci.org/totuworld/node-xlsx-json)

xlsx file을 json으로 변경하는 모듈이다.

rahil471의 모듈을 포크하여 정규표현식으로 특정 컬럼을 제외할 수 있는 옵션(exceptionColumn)을 추가하였다.

## Install

```
  npm install xlsx-to-json-yo
```

## Usage

```javascript
  xlsxj = require("xlsx-to-json");
  xlsxj({
    input: "sample.xlsx", 
    output: "output.json",
    exceptionColumn: "^_" //제외하고자하는 컬럼을 정규표현식으로 나타내면 된다. //예제는 언더바(_)바로 시작하는 컬럼을 제외하게 된다. 
  }, function(err, result) {
    if(err) {
      console.error(err);
    }else {
      console.log(result);
    }
  });
```

### Specifying a target sheet

You can optionally provide a sheet name to extract from that sheet

```javascript
  xlsxj = require("xlsx-to-json");
  xlsxj({
    input: "sample.xlsx", 
    output: "output.json",
    sheet: "tags"
  }, function(err, result) {
    if(err) {
      console.error(err);
    }else {
      console.log(result);
    }
  });
```

In config object, you have to enter an input path. But If you don't want to output any file you can set to `null`.

## License

MIT [@chilijung](http://github.com/chilijung)


