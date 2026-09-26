# PomoTodo

Pomodoro 番茄钟 + Todo list（Windows 托盘程序，.NET Framework 4.x，单文件 `PomoTodo.cs`）。

- 直接使用：`release/PomoTodo.exe`（当前版本 **v1.4.3**，窗口标题和 exe 文件属性里都会显示版本号）
- 自己编译：双击 `build.bat`（用 Windows 自带的 `csc.exe`）

## 版本历史

| 版本 | commit（在 `main` 上） | 说明 |
|---|---|---|
| v1.3.0 | `ca6388c` | 原始源码（baseline，只有源码没有 exe） |
| v1.4.0 | `92d769b`（PR #1） | 每台电脑一个文件的同步，并合并 Excel 记录。⚠️ exe 在 Windows 上报 `MissingMethodException`，打不开 |
| v1.4.1 | `aee81d7` | 出错时弹窗提示并写入 `crash.txt`。⚠️ 仍然有 `MissingMethodException` |
| v1.4.2 | `c6e093e`（PR #2） | 修复 `MissingMethodException`，第一个能在 Windows 上打开的 v1.4 版本 |
| v1.4.3 | `97c3ba0`（PR #3） | 窗口标题和文件属性显示版本号，功能和 v1.4.2 相同 |
| v1.5.0-beta | 测试版（PR 分支） | 任务下方新增可折叠的“已完成”区域：完成的任务自动移进去，默认折叠，可以按任务名或日期搜索；所有任务显示创建时间和完成时间 |

## v1.4 多电脑同步（OneDrive / Google Drive）

- 每台电脑只写自己的文件：`todos@电脑名.tsv`、`records@电脑名.tsv`、`deleted@电脑名.tsv`，
  两台电脑永远不写同一个文件 → 不再产生 `records-DESKTOP-XXX-2.tsv` 这类 conflict copies。
- 打开 / 每 20 秒同步时，自动合并文件夹里所有电脑的数据（包括旧的 `todos.tsv`、`records.tsv` 和旧的 conflict copies）：
  - Todo：同一个任务取最新修改的版本
  - 番茄记录：全部合并、去重
  - 删除：任何一台电脑删掉的任务 / 记录，在所有电脑上都删掉
- 已合并的旧文件移到同步文件夹里的 `old_files\`（不删除）。

## v1.4 Excel 记录（PomoTodo_Log.xlsx）

- 更新 Excel 之前，先把 Excel 本身和它的 conflict copies
  （`PomoTodo_Log-DESKTOP-XXX.xlsx`、`PomoTodo_Log (1).xlsx`）里的番茄记录合并回数据 →
  以前只留在 Excel 里的记录会恢复，重写 Excel 不会再丢历史。
- 内容没有变化时不重写 Excel → 两台电脑不再反复覆盖同一个文件。
- 合并过的 Excel conflict copies 移到 `old_files\`。

**两台电脑都要换成 v1.4。**

## 测试

```
PomoTodo.exe --synctest <空文件夹>    # 模拟两台电脑 + 旧 conflict copies + Excel 合并
PomoTodo.exe --selftest <file.xlsx>
```
