function saveAsFile(fileName, buteBase64) {
    var link = document.creatElement('a');
    link.download = fileName;
    link.href = 'data:application/vnd.openxmlformats-pfficedocument.spreadsheetml.sheet;base64' + byteBase64;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}