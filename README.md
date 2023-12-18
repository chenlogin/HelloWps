# wps加载项 项目

- 安装、更新 wps
```
npm install -g wpsjs
npm update -g wpsjs
```

- 创建项目、更新wps工具包
```
wpsjs create HelloWps 
npm update --save-dev wps-jsapi
```

- wpsjs工具包自动启动wps并加载HelloWps这个加载项
```
wpsjs debug
```

# HelloWps
- 文档
    - https://open.wps.cn/docs/client/wpsLoad
    - WPS集成模式 / WPS加载项开发
    - WPS基础接口
- wps加载项实现 对话框、任务窗格
    - src/main.js
- 浏览器唤起wps demo
    - http://127.0.0.1:3889/.debugTemp/systemdemo.html