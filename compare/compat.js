/* Edge / 旧 Chromium 补丁：IE 模式提示、arrayBuffer、Promise.finally 等 */
(function () {
  var ua = navigator.userAgent || "";
  var ieMode = !!document.documentMode || /MSIE |Trident\//.test(ua);
  var tooOld = ieMode || typeof Promise === "undefined" || typeof Map === "undefined";

  function showBanner() {
    var el = document.getElementById("oldBrowser");
    if (el) el.hidden = false;
  }

  if (tooOld) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", showBanner);
    } else {
      showBanner();
    }
  }

  if (window.NodeList && !NodeList.prototype.forEach) {
    NodeList.prototype.forEach = Array.prototype.forEach;
  }
  if (window.Element && !Element.prototype.matches) {
    Element.prototype.matches =
      Element.prototype.msMatchesSelector || Element.prototype.webkitMatchesSelector;
  }
  if (window.Element && !Element.prototype.closest) {
    Element.prototype.closest = function (sel) {
      var el = this;
      while (el && el.nodeType === 1) {
        if (el.matches(sel)) return el;
        el = el.parentElement;
      }
      return null;
    };
  }
  if (window.Blob && !Blob.prototype.arrayBuffer) {
    Blob.prototype.arrayBuffer = function () {
      var blob = this;
      return new Promise(function (resolve, reject) {
        var fr = new FileReader();
        fr.onload = function () {
          resolve(fr.result);
        };
        fr.onerror = function () {
          reject(fr.error || new Error("读取文件失败"));
        };
        fr.readAsArrayBuffer(blob);
      });
    };
  }
  if (window.Promise && !Promise.prototype.finally) {
    Promise.prototype.finally = function (cb) {
      return this.then(
        function (v) {
          return Promise.resolve(cb()).then(function () {
            return v;
          });
        },
        function (e) {
          return Promise.resolve(cb()).then(function () {
            throw e;
          });
        }
      );
    };
  }
})();
