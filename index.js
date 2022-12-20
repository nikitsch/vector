'use strict';
const excelToJson = require('convert-excel-to-json');
const fs = require('node:fs')

const result = excelToJson({
  sourceFile: 'TAGS-749.xls',
  columnToKey: {
    '*': '{{columnHeader}}'
  }
});

const PRIORITY = ['PV1', 'PV2', 'PV3', 'SAC1', 'SAC2', 'T1', 'QF1', 'SF1', 'SF2', 'SF3', 'SF4', 'SF5', 'SF6', 'SF7', 'SF8', 'SF9', 'SF10', 'SF11', 'SF12', 'KL1', 'KL2', 'FU', 'XT1', 'XT2', 'N', 'EG1', 'PE']
const NAMETERMINALS = "XT"

const MACHINES = Object.values(result)[0]

MACHINES.shift()

let newObjs = MACHINES.map(machine => {
  let keys = Object.keys(machine)
  return {
    ABR: machine.ABR,
    addresses: keys.reduce((sum, key) => {
      if (key.startsWith('ADDRES')) {
        return sum.concat({
          name: `${machine.ABR}:${key.split(':')[1]}`,
          value: machine[key]
        })
      }
      return sum
    }, [])
  }
})

newObjs = newObjs.map(el => {
  return {
    ...el, addresses: splitAddresses(el.addresses)
  }
})

function splitAddresses(addresses) {
  return addresses.reduce((sum, address) => {
    if (address.value.includes(', ')) {
      return sum.concat(address.value.split(', ')
        .map(point => ({
          ...address, value: point
        }))
      )
    }
    return sum.concat(address)
  }, [])
}


newObjs.sort((a, b) => PRIORITY.indexOf(a['ABR']) < PRIORITY.indexOf(b['ABR']) ? -1 : 1);
// console.log(JSON.stringify(newObjs, null, 2))

let addressArrays = newObjs.map(e => e.addresses)
let raisedUpObjects = addressArrays.flat()

let arrAddressesConsecutively = []

for (let [_, item] of Object.entries(raisedUpObjects)) {
  arrAddressesConsecutively.push(item.name, item.value)
}

const creatingAndSortingPairsArr = [];

let sortingByPriority= function (a, b) {
  if (PRIORITY.indexOf(a.split(':')[0]) < PRIORITY.indexOf(b.split(':')[0])) return -1
  if (PRIORITY.indexOf(a.split(':')[0]) > PRIORITY.indexOf(b.split(':')[0])) return 1
  let equals = (a.split(':')[1] < b.split(':')[1]) ? -1 : 1
  return equals
}

for (let i = 0; i < arrAddressesConsecutively.length; i += 2) {
  creatingAndSortingPairsArr.push(arrAddressesConsecutively.slice(i, i + 2).sort(sortingByPriority));
}

let createOneArrayString = creatingAndSortingPairsArr.join(' ').split(' ')

let successful = [],
  errors = [],
  unforeseenOutcome = []

let index = 0;
while (index < createOneArrayString.length) {
  let value = createOneArrayString[index];
  if (successful.includes(value) || errors.includes(value) || unforeseenOutcome.includes(value)) {
    index += 1
    continue;
  }
  let duplicatedValues = createOneArrayString.filter((item) => item === value);

  if (duplicatedValues.length == 1) {
    errors.push(value)
  } else if (duplicatedValues.length == 2) {
    successful.push(value)
  } else {
    unforeseenOutcome.push(value)
  }
  index += 1;
}

let tags = successful.map(addressSeparation => {
  let twoTags = addressSeparation.split(',')
  let firstTwoCharacters = twoTags.map(one => one.slice(0, 2))
  let lastCharacters = twoTags.map(one => one.slice(-1))

  let sortingByApostrophe = function (a) {
    return a.includes("'") ? 1 : -1
  }

  if (firstTwoCharacters.every(elem => elem == NAMETERMINALS) && lastCharacters.every(elem => elem == "'") ||
  firstTwoCharacters.every(elem => elem == NAMETERMINALS) && lastCharacters.every(elem => elem !== "'")) {
    return [twoTags.join(' '), twoTags.reverse().join(' ')]
  }
  if (firstTwoCharacters.every(elem => elem == NAMETERMINALS) && lastCharacters.indexOf("'") !== -1) {
    let apostrEnd = twoTags.slice().sort(sortingByApostrophe).join(' ')
    return [apostrEnd, apostrEnd]
  }
  if (firstTwoCharacters.indexOf(NAMETERMINALS) !== -1 && lastCharacters.indexOf("'") !== -1) {
    return [twoTags.slice().sort(sortingByApostrophe).join(' '), twoTags[lastCharacters.indexOf("'")]]
  }
  if (firstTwoCharacters.indexOf(NAMETERMINALS) !== -1) {
    return [twoTags.slice().sort((a) => a.includes(NAMETERMINALS) ? -1 : 1).join(' '), twoTags[firstTwoCharacters.indexOf(NAMETERMINALS)]]
  }
  return twoTags
})

fs.writeFileSync('convertedToExcel.json', JSON.stringify(tags))

// console.log('Бирки---------------------------------------------------------');
// tags.flat().map(e => console.log(e))
// console.log('--------------------------------------------------------------');
// console.log('Ошибки', errors);
// console.log('Неизвестная ошибка', unforeseenOutcome);//Больше 3 подключений это шина/Выводить надпись исправить и перезагрузить скрипт
