# 发布仓库指南

注册并登录自己的 GitHub 账号，在网页创建空仓库。先确认本机已安装 Git。
以下命令在项目目录执行，USER 应替换为自己的用户名。

```sh
git init
git add .
git status
```

提交前逐项检查暂存文件。合成演示账单和演示报告可以公开；真实账单、报名简介和私人联系方式不应进入公共仓库。默认忽略个人简介，通过私人报名包提供。

```sh
git commit -m "Add bill analyzer"
git branch -M main
git remote add origin https://github.com/USER/cpoe-wechat-bill-analyzer.git
git push -u origin main
```

登录和授权由账号所有者完成。推送后打开仓库，检查 README、文件和链接。填写实际仓库地址。

运行 `python tools/preflight.py` 做开发检查；报名包填完后执行 `python tools/preflight.py --submission`。两种检查都不能证明没有隐私信息，应人工核对。

.gitignore 只影响未跟踪文件，打 zip 时也不会自动应用它。公开前检查文件及历史提交。误发布真实数据后，删除当前文件不能清除历史，应先限制访问，再处理历史记录。
