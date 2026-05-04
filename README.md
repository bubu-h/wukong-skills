# wukong-skills

大师兄（@bubu-h）的 Claude Code Skill 集合，以 [Claude Code Plugin Marketplace](https://docs.claude.com/en/docs/claude-code/plugins) 形式发布。

## 当前包含的 Skill

| Skill | 说明 | 版本 |
|---|---|---|
| [credit-report-formatter](./credit-report-formatter/) | 将 markdown 格式的授信报告转换为标准格式的 docx 文件 | 0.1.3 |

## 安装方法

在 Claude Code 中执行以下两条命令：

```
/plugin marketplace add bubu-h/wukong-skills
/plugin install credit-report-formatter@wukong-skills
```

第一条命令把本仓库注册为 marketplace，**只需要执行一次**。
第二条命令安装具体的 skill，未来本仓库新增 skill 时，只需执行 `/plugin install <new-skill>@wukong-skills` 即可。

## 升级

```
/plugin update credit-report-formatter@wukong-skills
```

## 卸载

```
/plugin uninstall credit-report-formatter@wukong-skills
```

## 各 Skill 的详细使用说明

请进入对应子目录查看 README，例如：

- [credit-report-formatter/README.md](./credit-report-formatter/README.md)

## 贡献

本仓库主要供个人使用与少数朋友共享。如有问题或建议欢迎提 issue。

## License

MIT
