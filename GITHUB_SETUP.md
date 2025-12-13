# GitHub 发布指南

## 步骤 1：创建 GitHub 账户（如果还没有）

1. 访问 [GitHub 官网](https://github.com)
2. 点击右上角的 "Sign up" 按钮
3. 填写用户名、邮箱和密码
4. 完成邮箱验证

## 步骤 2：在 GitHub 上创建新仓库

1. 登录 GitHub 后，点击右上角的 "+" 号，选择 "New repository"
2. 填写仓库信息：
   - **Repository name**: `InvoiceAuto`（或你喜欢的名字）
   - **Description**: `自动化发票处理工具`（可选）
   - **Visibility**: 选择 Public（公开）或 Private（私有）
   - **⚠️ 重要**：**不要**勾选 "Initialize this repository with a README"（我们已经有了）
   - **不要**添加 .gitignore 或 license（我们已经有了）
3. 点击 "Create repository" 按钮

## 步骤 3：获取仓库地址

创建仓库后，GitHub 会显示仓库页面。你会看到类似这样的地址：
- HTTPS: `https://github.com/你的用户名/InvoiceAuto.git`
- SSH: `git@github.com:你的用户名/InvoiceAuto.git`

**复制 HTTPS 地址**（推荐使用 HTTPS）

## 步骤 4：配置远程仓库并推送代码

创建好仓库后，告诉我你的 GitHub 用户名和仓库名称，我会帮你配置并推送代码。

或者，你也可以手动执行以下命令（将 `你的用户名` 和 `仓库名` 替换为实际值）：

```bash
# 添加远程仓库
git remote add origin https://github.com/你的用户名/仓库名.git

# 推送代码到 GitHub
git branch -M main
git push -u origin main
```

## 步骤 5：更新 Git 用户信息（可选但推荐）

推送代码前，建议更新 Git 用户信息为你的 GitHub 信息：

```bash
git config --global user.name "你的GitHub用户名"
git config --global user.email "你的GitHub邮箱"
```

## 注意事项

1. **敏感信息**：你的代码中包含邮箱账号、密码和 API Key。在推送到 GitHub 前，建议：
   - 将这些信息移到配置文件（如 `config.ini`）
   - 将配置文件添加到 `.gitignore`
   - 或使用环境变量

2. **首次推送**：如果使用 HTTPS，GitHub 可能会要求你输入用户名和密码（或 Personal Access Token）

3. **Personal Access Token**：如果使用 HTTPS 推送，GitHub 现在要求使用 Personal Access Token 而不是密码：
   - 访问：Settings → Developer settings → Personal access tokens → Tokens (classic)
   - 生成新 token，勾选 `repo` 权限
   - 使用 token 作为密码

## 完成后

推送成功后，你的代码就会出现在 GitHub 上了！你可以：
- 在 GitHub 网页上查看代码
- 分享仓库链接给他人
- 继续使用 Git 进行版本控制

