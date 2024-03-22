// getting file

function load_ui() {
    const i = document.createElement("input");
    i.type = "file";
    document.body.appendChild(i);
    i.onchange = (e) => {
      const file = e.target.files[0];
      window.file = file;
      document.body.removeChild(i);
      delete i;
    };
  }
  
  async function getDatafromExcel(file) {
    const load_script = (uri) =>
      new Promise((res, rej) =>
        fetch(uri)
          .then((r) => r.text())
          .then(eval)
          .then(res)
          .catch(rej)
      );
    const libs = [
      "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js",
      "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js",
    ];
    const data_from_file = () =>
      new Promise((resolve, reject) => {
        reader = new FileReader();
        reader.onload = (e) => {
          const ext = file.name.split(".").pop().toUpperCase();
          const lib = window[ext];
          const workbook = lib.read(e.target.result, { type: "binary" });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const data = lib.utils.sheet_to_json(sheet);
          delete window.XLSX;
          delete window.XLS;
          delete window.file;
          delete window[ext];
          delete reader;
          resolve(data);
        };
        reader.readAsBinaryString(file);
      });
    return Promise.allSettled(libs.map(load_script)).then(data_from_file);
  }
  