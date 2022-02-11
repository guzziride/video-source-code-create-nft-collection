const basePath = process.cwd();
const fs = require("fs");
const jsonDir = `${basePath}/build/json`;
jsonFile = `${jsonDir}/_metadata.json`

let excel = require('excel4node');
let workbook = new excel.Workbook();
let tallyWorksheet = workbook.addWorksheet('Tally');
let individualWorksheet = workbook.addWorksheet('Individual');
let sortedWorksheet = workbook.addWorksheet('Individual_Sorted');

  // read the JSON metadata file and parse it
  const data = fs.readFileSync(jsonFile, {encoding:'utf8', flag:'r'});
  const nfts = JSON.parse(data);

  let nftTotal = 0;
  let tally = {
    "1 BGS" : {},
    "2 SHADE" : {},
    "3 BODY COLORS" : {},
    "4 SPARKLES" : {},
    "5 LIGHTS" : {},
    "6 EYES" : {},
    "7 GLASS" : {},
    "8 TIRES" : {},
    "9 RIMS" : {},
    "10 EXH" : {},
    "11 LINE" : {},
    "12 SKINS" : {},
    "13 SPOILERS" : {},
    "14 MOUTH" : {},
    "15 PERSONALITY" : {},
    "16 FX" : {},
  };


  // Compute the tally for each value for each trait
  let i = 0;
  while (nfts[i] != null) {
     let nft = nfts[i];
    let attributes = nft["attributes"];
    attributes.forEach(element => {
      let traitType = element["trait_type"];
      let value = element["value"];
      if(tally[traitType].hasOwnProperty(value)) {
        // increment the tally for that trait value (e.g. [RIMS][Rims6Pink])
        tally[traitType][value]++;
      } else {
        // initialize the tally for that trait value
        tally[traitType][value] = 1;
      }
    });
    i++;
    nftTotal = i;
  };
  console.log('nft total: ', nftTotal);
  console.log('tally: ', tally);


let style = workbook.createStyle({
  font: {
    color: '#FF0800',
    size: 12
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});

tallyWorksheet.cell(1,1).string('TRAIT');
tallyWorksheet.cell(1,2).string('VALUE');
tallyWorksheet.cell(1,3).string('COUNT');
tallyWorksheet.cell(1,4).string('RARITY');

// Compute the rarity for each trait
let rarityCounts = {};
let row = 2;
for (traitType in tally) {
  if (tally.hasOwnProperty(traitType)) {
    tallyWorksheet.cell(row,1).string(traitType); row++;
    for (value in tally[traitType]) {
      if (tally[traitType].hasOwnProperty(value)) {
        tallyWorksheet.cell(row,2).string(value);
        let count = tally[traitType][value];
        let rarity = nftTotal/count;
        tallyWorksheet.cell(row,3).number(count);
        tallyWorksheet.cell(row,4).number(rarity);
        rarityCounts[value] = rarity;
        console.log('Trait value: ', value, '  rarity: ', rarity);
        row++;
      }
    }
  }
  row++;
}

// Now compute the rarity for each NFT
i = 0;
while (nfts[i] != null) {
  let nft = nfts[i];
  row = i+1;
  individualWorksheet.cell(row,1).number(row);
  sortedWorksheet.cell(row,1).number(row);
  individualWorksheet.cell(row,2).string(nft["name"]);
  sortedWorksheet.cell(row,2).string(nft["name"]);
  let rarityScore = 0;
  let attributes = nft["attributes"];
  attributes.forEach(element => {
//    let traitType = element["trait_type"];
    let value = element["value"];
    rarityScore = rarityScore + rarityCounts[value];
  });
  individualWorksheet.cell(row,3).number(rarityScore);
  sortedWorksheet.cell(row,3).number(rarityScore);
  //console.log(nft["name"], ': ', rarityScore);
  i++;
};

workbook.write('Rarity.xlsx');