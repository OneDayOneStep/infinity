importScripts("./xlsx.mini.min.js");
self.onmessage = ev => {
  const excelData = XLSX.read(ev.data);
  postMessage(excelData);
};
