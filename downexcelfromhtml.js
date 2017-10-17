// 前端下载文件



(function (global, factory) {

  console.log(global, factory)
  typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory() :
    typeof define === 'function' && define.amd ? define(factory) : (global.download = factory());

})(this, function () {




  var idTmr

  function getExplorer() {
    var explorer = window.navigator.userAgent
    if (explorer.indexOf('MSIE') >= 0) {
      return 'ie'
    } else if (explorer.indexOf('Firefox') >= 0) {
      return 'Firefox'
    } else if (explorer.indexOf('Chrome') >= 0) {
      return 'Chrome'
    } else if (explorer.indexOf('Opera') >= 0) {
      return 'Opera'
    } else if (explorer.indexOf('Safari') >= 0) {
      return 'Safari'
    }
  }


  function download(slt, filename) {
    if (getExplorer() === 'ie') {
      var curTbl = document.querySelector(slt)
      try {
        var oXL = new ActiveXObject('Excel.Application')
      } catch (err) {
        alert('当前浏览器无法下载列表数据，请使用chrome或者firefox进行下载！')
        return
      }
      var oWB = oXL.Workbooks.Add()
      var xlsheet = oWB.Worksheets(1)
      var sel = document.body.createTextRange()
      sel.moveToElementText(curTbl)
      sel.select()
      sel.execCommand('Copy')
      xlsheet.Paste()
      oXL.Visible = true
      try {
        var fname = oXL.Application.GetSaveAsFilename('Excel.xls',
          'Excel Spreadsheets (*.xls),   *.xls')
      } catch (e) {
        print('Nested catch caught ' + e)
      } finally {
        oWB.SaveAs(fname)
        oWB.Close(savechanges = false)
        oXL.Quit()
        oXL = null
        idTmr = window.setInterval('Cleanup();', 1)
      }
    } else {
      tableToExcel(slt, filename)
    }
  }
  // cleanup
  function Cleanup() {
    window.clearInterval(idTmr)
    CollectGarbage()
  }
  var tableToExcel = (function () {
    var uri = 'data:application/vnd.ms-excel;base64,',
      template = '<html><head><meta charset="UTF-8"></head><body><table>{table}</table></body></html>',
      base64 = function (s) {
        return window.btoa(unescape(encodeURIComponent(s)))
      },
      format = function (s, c) {
        return s.replace(/{(\w+)}/g, function (m, p) {
          return c[p]
        })
      }
    return function (table, name) {
      if (!table.nodeType) table = document.querySelector(table)
      var ctx = {
        worksheet: name || 'Worksheet',
        table: table.innerHTML
      }
      // console.log('表格内容：：：', ctx.table)
      try {
        window.location.href = uri + base64(format(template, ctx))
      } catch (err) {
        alert('当前浏览器无法下载列表数据，请使用chrome或者firefox进行下载！')
        return
      }
    }
  })()

  return download

})
