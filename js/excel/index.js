

// "https://raw.githubusercontent.com/prajwalgangawane/my_utils_js/main/js/excel/index.js"


// Demo purpose only
/**
@author prajwalgangawane
@function load_ui
This function dynamically creates an input element of type file,
appends it to the document body, and sets up an onchange event listener
to handle file selection. Once a file is selected, it is stored in the
window object as 'file', and the input element is removed from the DOM.
*/
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

/**
 * @author prajwalgangawane
 * @function getDatafromExcel
 * @param {File} file - The Excel file from which data will be extracted
 * @returns {Promise<Array<Object>>} - A promise that resolves with an array of objects representing the data from the Excel file
 * This function getDatafromExcel asynchronously extracts data from an Excel file.
 * It loads necessary libraries dynamically, processes the file, and returns a promise that resolves with an array of objects representing the data from the Excel file.
 */

async function getDatafromExcel(file) {
    const load_script = (uri) =>
        new Promise((res, rej) =>
            fetch(uri)
                .then((r) => r.text())
                .then(eval)
                .then(res)
                .catch(rej)
        );
    // const libs = [
    //   "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js",
    //   "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js",
    // ];
    const libs = [
        "https://raw.githubusercontent.com/prajwalgangawane/my_utils_js/main/js/excel/deps/jszip.js",
        "https://raw.githubusercontent.com/prajwalgangawane/my_utils_js/main/js/excel/deps/xlsx.js",
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
                delete getDatafromExcel;
                delete load_ui;
                resolve(data);
            };
            reader.readAsBinaryString(file);
        });
    return Promise.allSettled(libs.map(load_script)).then(data_from_file);
}
