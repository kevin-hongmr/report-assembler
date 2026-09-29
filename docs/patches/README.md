# Windows 基线（main 分支）修复补丁：关系条目命名空间

## 问题

`app/core/assembler.py` 的 `_patch_package()` 在补写图片 / OLE 关系条目与 Content_Types 声明时，
把 `.rels` 部件的子元素建到了错误的 XML 命名空间：

| 用途 | 命名空间 |
| ---- | -------- |
| `r:id` 属性、关系 `Type` 取值 | `http://schemas.openxmlformats.org/officeDocument/2006/relationships`（代码里的 `R_NS`） |
| **`.rels` 部件自身（根元素与其子元素）** | `http://schemas.openxmlformats.org/package/2006/relationships` |

误用 `R_NS` 后，`.rels` 里会出现类似下面的条目：

```xml
<ns0:Relationship xmlns:ns0="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  Id="rId1001" Type=".../image" Target="media/asm_1001.wmf"/>
```

解析器会把它当成**外来元素直接忽略** → `word/document.xml` 里的 `r:id` 悬空 →
**含图片 / 嵌入对象的成品在 Microsoft Word、LibreOffice 下会丢图，甚至被判定为文件损坏。**

WPS 对关系命名空间较宽容，所以这个问题在 Windows 上一贯不报错、不被察觉；
麒麟侧用 LibreOffice 预览时才会暴露（`Error: source file could not be loaded`）。

**实测影响**：旧版本产出的 `测试2/汇编结果（测试2）.docx` 在 LibreOffice 下完全无法加载；
修复后同一文件可正常转出 86 页 PDF，图片绘制在第 54/76/80 页。

## 这个补丁改什么

只有 `app/core/assembler.py` 一个文件，共 5 处：

1. 新增两个常量 `PKG_R_NS`（package/2006/relationships）与 `CT_NS`（package/2006/content-types）；
2. 两处新建 `<Relationship>`：`R_NS` → `PKG_R_NS`（媒体/OLE 段、补 numbering 关系段）；
3. 两处新建 Content_Types 的 `Default` / `Override`：由硬编码字符串改为 `CT_NS`。

**不涉及任何国产系统 / LibreOffice 相关代码**，可安全应用到 `main`。

## 怎么用

### 方式一（推荐）：直接替换文件

`assembler.py.main-fixed` 就是**打好补丁的完整文件**。把它复制到 Windows 工作副本的
`app\core\` 下、改名为 `assembler.py`（覆盖原文件）即可。

已验证：它与应用补丁后的 `main` 版**逐字节一致**（MD5 `a3df91112d7b9ee47c550fc4f5967c4e`），
也与两个国产系统离线包内的同名文件一致。

> 覆盖前请先备份原文件；改完建议按下面的"验证"跑一遍。

### 方式二：手工改两行（不想替换整个文件时）

在 `app/core/assembler.py` 中：

```diff
 # 关系（relationships）命名空间，用于图片/编号等外部引用
 R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
+PKG_R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
+CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
```

然后把**两处** `etree.SubElement(rel_root, "{%s}Relationship" % R_NS)` 改为
`etree.SubElement(rel_root, "{%s}Relationship" % PKG_R_NS)`（分别在 `_patch_package()` 的
媒体/OLE 段与补 numbering 关系段），另外两处 Content_Types 的 `SubElement` 改用 `CT_NS`。

等价的可机读补丁见 `rels-namespace.patch`（`git apply` 可直接用）。

## 验证

改完后出包，然后用**含 .doc/.wps（带图片或嵌入对象）**的多份材料汇编一份成品，
检查两件事：

1. 成品能用 Microsoft Word 打开且图片完整显示；
2. 成品能被 LibreOffice 打开（`soffice --headless --convert-to pdf 成品.docx` 能成功）。

也可直接跑仓库里的离线回归测试（纯 Python，不依赖 WPS / LibreOffice）：

```bash
python tests/test_opc_package_integrity.py
```

它会在缺陷复现时报出"关系条目命名空间错误"与"document.xml 悬空引用"。

## 已经分发出的成品怎么办

旧版本产出的、含图片或嵌入对象的成品同样带这个缺陷。两条路：

- **重新汇编**（推荐，可一并对齐其它格式问题）；
- 用 `tools/fix_docx_rels_namespace.py` 就地修复（只治命名空间这一种缺陷，其余部件不动）。
