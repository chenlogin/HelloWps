
function openOfficeFileFromSystemDemo(param){
    let jsonObj = (typeof(param)=='string' ? JSON.parse(param) : param)
    alert("从业务系统传过来的参数为：" + JSON.stringify(jsonObj))
    
    //打开指定文件
    wps.WppApplication().Presentations.OpenFromUrl('http://127.0.0.1:3889/.debugTemp/wppDemo.pptx')
    //插入新页、添加背景、添加超链、播放幻灯片
    insertImage()

    return {wps加载项项返回: jsonObj.filepath + ", 这个地址给的不正确"}
}

function insertImage() {
    try {
      const activeDoc = window.wps.Application.ActivePresentation
      const allWindow = window.wps.Application.Windows

      let activeWindow = allWindow.Item(1)
      for (let i = 1; i <= allWindow.Count; i++) {
        const item = window.wps.Application.Windows.Item(i)
        if (item.Caption === activeDoc.Name) {
          activeWindow = window.wps.Application.Windows.Item(i)
          break
        }
      }
      const slideRange = activeWindow.Selection.SlideRange
      let pageIndex = 0 // 当前选中幻灯片的下标，从1开始，0表示未选中
      if (slideRange && slideRange.Count) {
        pageIndex = slideRange.Item(slideRange.Count).SlideIndex
      } else {
        // 如果没找到就取当前文档的幻灯片最大个数
        pageIndex = activeDoc.Slides.Count || 0
      }
      // 在当前幻灯片后添加新幻灯片
      activeDoc.Slides.Add(pageIndex + 1, 12)
      const newSlide = activeDoc.Slides.Item(pageIndex + 1)
      // 在新幻灯片中添加背景图片
      const pageSetup = activeDoc.PageSetup
      const coverPath = "C:\\workspace\\HelloWps\\public\\images\\bg.jpg"
      newSlide.Shapes.AddPicture(coverPath, false, true, 0, 0, pageSetup.SlideWidth, pageSetup.SlideHeight)
      // 在新幻灯片中添加左上角的icon
      const iconPath = "C:\\workspace\\HelloWps\\public\\images\\icon.png";
      const iconShape = newSlide.Shapes.AddPicture(iconPath, false, true, 0, 16, 168, 56)
      const iconPicture = iconShape.ActionSettings.Item(1)
      // 给icon添加超链接
      iconPicture.Hyperlink.Action = 7
      iconPicture.Hyperlink.Address = "https://www.baidu.com"

      // 播放幻灯片
      activeDoc.SlideShowSettings.Run();
    } catch (e) {
      window.wps.alert('添加失败，请重试')
    }
    window.close()
  }

function InvokeFromSystemDemo(param){
    let jsonObj = (typeof(param)=='string' ? JSON.parse(param) : param)
    let handleInfo = jsonObj.Index
    switch (handleInfo){
        case "getDocumentName":{
            let docName = ""
            if (wps.WppApplication().ActivePresentation){
                docName = wps.WppApplication().ActivePresentation.Name
            }

            return {当前打开的文件名为:docName}
        }

        case "newDocument":{
            let newDocName=""
            let doc = wps.WppApplication().Presentations.Add()
            newDocName = doc.Name
            
            return {操作结果:"新建文档成功，文档名为：" + newDocName}
        }

        case "OpenFile":{
            let filePath = jsonObj.filepath
            wps.WppApplication().Presentations.OpenFromUrl(filePath)
            return {操作结果:"打开文件成功"}
        }
    }

    return {其它xxx:""}
}

export default{
    openOfficeFileFromSystemDemo,
    InvokeFromSystemDemo
}