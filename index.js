const express = require('express')
const app = express()
const port = 3000

global.groupdocs_merger_cloud = require("groupdocs-merger-cloud");
global.fs = require("fs");

global.clientId = "894d5361-5fcb-4ab2-ab91-af394e0a78aa";
global.clientSecret = "d19b0572a2298026850ca68d1d195419";
global.myStorage = "MergePPTX";
const config = new groupdocs_merger_cloud.Configuration(clientId, clientSecret);
config.apiBaseUrl = "https://api.groupdocs.cloud";

let fileApi = groupdocs_merger_cloud.FileApi.fromConfig(config);
// open multiple pptx files folder from your system drive.
let resourcesFolder = 'E:\\Tasks\\\MergingPPTX\\public\\';

app.get('/', (req, res) => {
  res.send('Hello World!')
})

const upload = () => {
  fs.readdir(resourcesFolder, (err, files) => {
    files.forEach(file => {
      // read files one by one
      fs.readFile(resourcesFolder + file, (err, fileStream) => {
        // create upload file request
        let request = new groupdocs_merger_cloud.UploadFileRequest(file, fileStream, myStorage);
        // upload file
        fileApi.uploadFile(request)
      });
    });
  });
}

const combine = async (path1, path2, mergedPath) => {
  let documentApi = groupdocs_merger_cloud.DocumentApi.fromKeys(clientId, clientSecret);
  
  // create first join item
  let item1 = new groupdocs_merger_cloud.JoinItem();
  item1.fileInfo = new groupdocs_merger_cloud.FileInfo();
  item1.fileInfo.filePath = path1;
  
  // create second join item
  let item2 = new groupdocs_merger_cloud.JoinItem();
  item2.fileInfo = new groupdocs_merger_cloud.FileInfo();
  item2.fileInfo.filePath = path2;
  
  // create join options
  let options = new groupdocs_merger_cloud.JoinOptions();
  options.joinItems = [item1, item2];
  options.outputPath = mergedPath;
  
  try {
    // Create join documents request
    let joinRequest = new groupdocs_merger_cloud.JoinRequest(options);
    let result = await documentApi.join(joinRequest);
  } 
  catch (err) {
    throw err;
  }
}

const download = () => {
  const fileApi = groupdocs_merger_cloud.FileApi.fromConfig(config);
    // create donwload file request
    let request = new groupdocs_merger_cloud.DownloadFileRequest("joined-file.pptx", myStorage);
    // download file and response type Stream
    fileApi.downloadFile(request)
      .then(function (response) {
          // save file in your system directory
          fs.writeFile(resourcesFolder + "joined-file.pptx", response, "binary", function (err) { });
          console.log("Expected response type is Stream: " + response.length);
      })
      .catch(function (error) {
          console.log("Error: " + error.message);
      });
}

app.listen(port, () => {
  upload()
  combine("Presentation1.pptx", "Presentation2.pptx", "joined-file.pptx")
    .then((res) => {
      download()
      console.log("Successfully combined powerpoint pptx files: ");
    })
    .catch((err) => {
      console.log("Error occurred while merging the PowerPoint files:", err);
    })

  console.log(`Example app listening on port ${port}`)
})
