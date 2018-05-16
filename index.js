// dotenv environment variables
require('dotenv').config()
const fs = require('fs');
const xlsx = require('xlsx');
const CognitiveServicesCredentials = require('ms-rest-azure').CognitiveServicesCredentials;
// Creating the Cognitive Services credentials
// This requires a key corresponding to the service being used (i.e. text-analytics, etc)
const credentials = new CognitiveServicesCredentials(process.env.COGNITIVE_SERVICES_CREDENTIALS);
const region = process.env.COGNITIVE_SERVICES_REGION;
const TextAnalyticsAPIClient = require('azure-cognitiveservices-textanalytics');

try{
  fs.mkdir('./output');
}catch(_ignored){
}

let client = new TextAnalyticsAPIClient(credentials, region);

const data = xlsx.readFile('./sample/SatisfactionSurvey.xlsx');
// Array of objects {comments: 'some comment', 'Completion time': '4/11/18 ...', Email: 'some.email@email.com', Name: 'Some name', recoment: "10", 'Start time': '4/11/18'}
const jsonData = xlsx.utils.sheet_to_json(data.Sheets[data.SheetNames[0]]);

const documents = jsonData.map((row, i) => {
  return { id: '' + i, text: row.comments };
})

Promise.all([
  client.sentiment({
    documents: documents
  }),
  client.keyPhrases({
    documents: documents
  })
]).then((res) => {
  const sentiments = res[0];
  const keyPhrases = res[1];

  //Map of id to Score
  const sentimentIdToValueMap = new Map();
  sentiments.documents.forEach((s) => {
    sentimentIdToValueMap.set(s.id, s.score);
  });
  // map of id to keyPhrases
  const keyPhrasesIdToValueMap = new Map();
  keyPhrases.documents.forEach((k) => {
    keyPhrasesIdToValueMap.set(k.id, k.keyPhrases);
  });

  // log possible errors in data
  sentiments.errors && sentiments.errors.lenght && console.log('Sentiments errors', JSON.stringify(sentiments.errors, null, 2));
  keyPhrases.errors && keyPhrases.errors.lenght && console.log('key phrase errors', JSON.stringify(keyPhrases.errors, null, 2));

  const outputJsonData = jsonData.map((row, i) => {
    const id = i + '';
    const score = sentimentIdToValueMap.get(id);
    const sentiment = score !== undefined ? (score > 0.5 ? 'positive' : 'negative') : 'Impossible to calculate';
    const keyPhrases = (keyPhrasesIdToValueMap.get(id) || []).join();

    return Object.assign(row, { sentimentScore: score, sentiment: sentiment, keyPhrases: keyPhrases });
  });

  var wb = { SheetNames: ['Sheet1'], Sheets: { 'Sheet1': xlsx.utils.json_to_sheet(outputJsonData) } };
  xlsx.writeFile(wb, './output/output.xlsx');
})

