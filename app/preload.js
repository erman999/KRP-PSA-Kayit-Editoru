// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
window.addEventListener('DOMContentLoaded', () => {
  // console.log("Ready.");

  var fs = require('fs');
  var path = require('path');
  var XLSXWriter = require('xlsx-writestream');
  const { shell } = require('electron');

  var headings = [
    "Dosya Adı",
    "Dosya Tarihi",
    "Seri No",
    "OP20 Pim-1 çakma kN",
    "Sonuç[1]",
    "OP20 Pim-2 çakma kN",
    "Sonuç[2]",
    "OP20 Tapa çakma kN",
    "Sonuç[3]",
    "OP30 Rulman çakma kN",
    "Sonuç[4]",
    "OP30 Rulman çakma sonu mm",
    "Sonuç[5]",
    "OP30 Kömür çakma kN",
    "Sonuç[6]",
    "OP30 Kömür çakma max. kN",
    "Sonuç[7]",
    "OP30 Kömür çakma sonu mm",
    "Sonuç[8]",
    "OP35 Kömür yükseklik mm",
    "Sonuç[9]",
    "OP40 Kasnak iç çap mm",
    "Sonuç[10]",
    "OP40 Kasnak iç çap ovalite mm",
    "Sonuç[11]",
    "OP40 Kasnak çakma kN",
    "Sonuç[12]",
    "OP40 Pervane iç çap mm",
    "Sonuç[13]",
    "OP40 Pervane iç çap ovalite mm",
    "Sonuç[14]",
    "OP40 Pervane çakma kN",
    "Sonuç[15]",
    "OP50 Boşta dönme kontrolü",
    "Sonuç[16]",
    "OP50 Kasnak yükseklik mm",
    "Sonuç[17]",
    "OP50 Kasnak dış çap ovalite mm",
    "Sonuç[18]",
    "OP50 Pervane yükseklik mm",
    "Sonuç[19]",
    "OP50 Pervane dış çap ovalite mm",
    "Sonuç[20]",
    "OP60 Sıkılık torku Nm",
    "Sonuç[21]",
    "OP70 Sızdırmazlık Pa",
    "Sonuç[22]",
    "OP80 Conta kulak ve kirlilik kamera kontrolü",
    "Sonuç[23]"
  ];

  var colHeadings = "";
  headings.forEach(function(data, index){
    colHeadings += data + "\t";
  });
  colHeadings += "\n";

  var fileList = [];
  var walkSync = function(dir, fileList) {
    var fs = fs || require('fs');
    var path = path || require('path');
    var files = fs.readdirSync(dir);
    fileList = fileList || [];
    files.forEach(function(file) {
      var joined_path = path.join(dir, file);
      // console.log(joined_path);
      if (fs.statSync(joined_path).isDirectory()) {
        fileList = walkSync(joined_path, fileList);
      }
      else {
        fileList.push(joined_path);
      }
    });
    return fileList;
  };


  function dateFormat(d){
    var _day = ("0" + d.getDate()).slice(-2);
    var _month = ("0"+(d.getMonth()+1)).slice(-2);
    var _year = d.getFullYear();
    var _hours = ("0" + d.getHours()).slice(-2);
    var _minutes = ("0" + d.getMinutes()).slice(-2);
    var _seconds = ("0" + d.getSeconds()).slice(-2);
    var datestring = _year +"-"+ _month +"-"+ _day +" "+ _hours +":"+ _minutes +":"+ _seconds;
    return datestring;
  }

  function fix_content(filePath, readedFile){

    // Encoding: ascii, base64, binary, hex, ucs2, utf8, latin1
    // var readedFile = fs.readFileSync(filePath, "binary");

    // Remove whitespace from both sides of the whole string
    readedFile = readedFile.trim();

    // Remove \r
    readedFile = readedFile.replace(/\r/g, '');

    // Convert missing values to tab
    readedFile = readedFile.replace(/\s\s\s\s\s\s\s\s\s\s\s/gm, '\tA_A');

    // Convert missing values to tab
    readedFile = readedFile.replace(/mm\s*OK/gm, 'mm\tA_A\tOK');

    // remove spaces before and after \t
    readedFile = readedFile.replace(/( {1,})+\t/g, '\t'); // removes spaces before tab
    readedFile = readedFile.replace(/\t+( {1,})/g, '\t'); // removes spaces after tab

    // remove spaces before and after \n
    readedFile = readedFile.replace(/\n+( {1,})/g, '\n'); // removes spaces before newline
    readedFile = readedFile.replace(/( {1,})+\n/g, '\n'); // removes spaces after newline

    // Remove operation names
    readedFile = readedFile.replace(/^OP([\s\S]*?)(\t\s|\t)/gm, '');

    // Convert '...', '???', 'xxx' values to tab
    readedFile = readedFile.replace(/(\.\.\.)/gm, 'A_A');
    readedFile = readedFile.replace(/(\?\?\?)/gm, 'A_A');
    readedFile = readedFile.replace(/(xxx)/gm, 'A_A');

    // Fix 3 space error
    readedFile = readedFile.replace(/\s\s\s/gm, '\t');

    // Change decimal separator (sometimes it may vary from country to country)
    // This changes decimal separator from comma to point
    readedFile = readedFile.replace(/(\d)(\,)(\d)/g, '$1.$3');

    // This changes decimal separator from point to comma
    // readedFile = readedFile.replace(/(\d)(\.)(\d)/g, '$1,$3');

    // Change \n (newline) with \t (tab)
    readedFile = readedFile.replace(/\n/g, '\t');

    // Read file stats
    stats = fs.statSync(filePath);

    // Add file last modified datetime (date format is dd-mm-yyyy hh-mm-ss)
    readedFile = dateFormat(stats.mtime) + "\t" + readedFile;

    // Add file name
    readedFile = path.basename(filePath) + "\t" + readedFile;

    // PSA data requires 48 columns fixTabularData
    var tabularColumnNum = 48;

    var count = (readedFile.match(/\t/g) || []).length;
    if (count < tabularColumnNum) {
      var loop = tabularColumnNum - count;
      for (var i = 0; i < loop; i++) {
        readedFile += "\tA_A";
      }
    }

    // Remove A_A strings removeA_A
    readedFile = readedFile.replace(/A_A/gm, '');

    // Add new line at the end of each loop
    readedFile += '\n';

    return readedFile;
  }

  function toObj(key, arr){
    var myObj = {};
    for (var i = 0; i < arr.length; i++) {
      if (isNaN(arr[i])) {
        myObj[key[i]] = arr[i];
      } else {
        myObj[key[i]] = parseFloat(arr[i]);
      }
    }
    return myObj;
  }

  const holder = document.getElementById('holder');
  const filename = document.getElementById('filename');
  const type = document.getElementById('type');
  const filepath = document.getElementById('filepath');
  const btn_start = document.getElementById('btn_start');
  const progress_bar = document.getElementById('progress_bar');
  const btn_reset = document.getElementById('btn_reset');
  const events = document.getElementById('events');
  const srcLink = document.getElementById('srcLink');
  var file_basename = 'output';

  btn_reset.addEventListener("click", function(){
    // console.log("Reset Clicked.");
    btn_start.disabled = false;
    file_basename = 'output';
    fileList = []; // Clear already found file paths
    progress_bar.innerHTML = '0%';
    progress_bar.style.width = '0%';
    holder.className = '';
    type.innerHTML = '';
    filename.innerHTML = '';
    filepath.innerHTML = '';
    events.innerHTML = 'Durum';
  });

  srcLink.addEventListener("click", function(){
    shell.openExternal('https://github.com/erman999');
  });

  // prevent default behavior from changing page on dropped file
  window.ondragover = function(e) { e.preventDefault(); return false };
  // NOTE: ondrop events WILL NOT WORK if you do not "preventDefault" in the ondragover event!!
  window.ondrop = function(e) { e.preventDefault(); return false };

  holder.ondragover = function () { this.className = 'hover'; return false; };
  holder.ondragleave = function () { this.className = ''; return false; };

  holder.ondrop = function (e) {
    e.preventDefault();

    btn_start.disabled = false;
    file_basename = 'output';
    fileList = []; // Clear already found file paths
    progress_bar.innerHTML = '0%';
    progress_bar.style.width = '0%';
    events.innerHTML = 'Dosyalar bulunuyor...';

    const transfer = e.dataTransfer.items[0];
    const entry = transfer.webkitGetAsEntry();
    const getFile = transfer.getAsFile();

    // console.log(getFile);
    // console.log("isFile:", entry.isFile);
    // console.log("isDirectory:", entry.isDirectory);
    // console.log("File name:", getFile.name);
    // console.log("File path:", getFile.path);

    file_basename = getFile.name;
    filename.innerHTML = file_basename;
    filepath.innerHTML = getFile.path;

    if (entry.isFile) {
      type.innerHTML = "Dosya";
      fileList.push(getFile.path);
    } else {
      type.innerHTML = "Klasör";
      walkSync(getFile.path, fileList);
    }

    events.innerHTML = 'Dosyalar bulundu. Başlamaya hazır.';

    return false;
  };

  btn_start.addEventListener("click", function(){
    // console.log("Start Clicked.");
    btn_start.disabled = true;

    var total = fileList.length;
    var counter = 0;

    if (fileList.length > 0) {
      var dir = './excel';
      if (!fs.existsSync(dir)){
        fs.mkdirSync(dir);
      }
      var writer = new XLSXWriter( 'excel/' + file_basename + '.xlsx', {} /* options */);
      writer.getReadStream().pipe(fs.createWriteStream( 'excel/' + file_basename + '.xlsx'));
    }


    fileList.forEach(function(item, index){

      fs.readFile(item, 'binary', function(err, data){
        if (err) throw err;
        // console.log(data); // Check what is read

        setTimeout(function(){
          var myData = fix_content(item, data);
          myData = myData.trim();
          myData = myData.split(/\t/);
          // console.log(myData);
          // console.log(headings);

          writer.addRow(toObj(headings, myData));

          counter++;
          var percent = Math.floor((counter/total)*1000)/10;
          // console.log(percent);
          events.innerHTML = 'İşleniyor: ' + path.basename(item);

          progress_bar.innerHTML = percent + '%';
          progress_bar.style.width = percent + '%';

          if (percent == 100) {
            events.innerHTML = 'Tamamlandı';
          }

          // console.log(index+1, total);

          if (index+1 == total) {
            setTimeout(function(){
              writer.finalize();
              alert('Dosya oluşturuldu: ' + file_basename + '.xlsx');
              var outputPath = path.join(path.dirname(__dirname), 'excel', file_basename + '.xlsx');
              // console.log(outputPath);
              shell.showItemInFolder(outputPath);
              btn_start.disabled = false;
            }, 500);
          }

        }, 0);

      });

    });

  });

});
