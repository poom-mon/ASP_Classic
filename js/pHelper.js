var loaderProgress = function (wording) {
  //   this.wording = wording;
    $("#lblLoader").innerHtml(wording);
};
loaderProgress.prototype.loader = function () {
    $("#processing-modal").modal("show")
};
loaderProgress.prototype.loadSuccess = function () {
    $("#processing-modal").modal("hide")
};

