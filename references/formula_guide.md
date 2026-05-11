# Excel/WPS 公式速查手册

## 逻辑函数

| 函数 | 语法 | 说明 | 示例 |
|------|------|------|------|
| IF | `IF(条件, 真值, 假值)` | 条件判断 | `IF(A1>60,"及格","不及格")` |
| AND | `AND(条件1, 条件2, ...)` | 多条件同时成立 | `AND(A1>60,B1="是")` |
| OR | `OR(条件1, 条件2, ...)` | 任一条件成立 | `OR(A1>90,B1="优秀")` |
| NOT | `NOT(条件)` | 逻辑取反 | `NOT(A1="")` |
| IFERROR | `IFERROR(公式, 替代值)` | 错误时返回替代值 | `IFERROR(A1/B1,0)` |

## 查找引用函数

| 函数 | 语法 | 说明 |
|------|------|------|
| VLOOKUP | `VLOOKUP(查找值, 表格范围, 返回列号, 匹配方式)` | 垂直查找（0=精确, 1=模糊） |
| XLOOKUP | `XLOOKUP(查找值, 查找列, 返回列)` | 新版万能查找（推荐） |
| INDEX | `INDEX(范围, 行号, 列号)` | 返回指定位置的值 |
| MATCH | `MATCH(查找值, 范围, 匹配方式)` | 返回相对位置 |
| HLOOKUP | `HLOOKUP(查找值, 表格范围, 返回行号, 匹配方式)` | 水平查找 |

## 文本处理函数

| 函数 | 语法 | 说明 |
|------|------|------|
| LEFT | `LEFT(文本, 字符数)` | 取左侧字符 |
| RIGHT | `RIGHT(文本, 字符数)` | 取右侧字符 |
| MID | `MID(文本, 开始位置, 字符数)` | 取中间字符 |
| LEN | `LEN(文本)` | 计算文本长度 |
| TEXTJOIN | `TEXTJOIN(分隔符, 忽略空, 范围1, ...)` | 合并文本 |
| CONCAT | `CONCAT(文本1, 文本2, ...)` | 连接文本 |
| TRIM | `TRIM(文本)` | 去除多余空格 |
| SUBSTITUTE | `SUBSTITUTE(文本, 旧, 新, 替换次数)` | 替换文本 |

## 日期函数

| 函数 | 语法 | 说明 |
|------|------|------|
| DATEDIF | `DATEDIF(开始日, 结束日, "单位")` | 计算日期间隔（Y/M/D） |
| TODAY | `TODAY()` | 当前日期 |
| NETWORKDAYS | `NETWORKDAYS(开始日, 结束日, 节假日)` | 工作日天数 |
| EOMONTH | `EOMONTH(日期, 月数)` | 月末日期 |
| YEAR/MONTH/DAY | `YEAR(日期)` | 提取年/月/日 |
| WEEKDAY | `WEEKDAY(日期, 返回类型)` | 星期几（2=周一1~周日7） |

## 数学计算函数

| 函数 | 语法 | 说明 |
|------|------|------|
| SUM | `SUM(范围)` | 求和 |
| SUMIF | `SUMIF(条件范围, 条件, 求和范围)` | 条件求和 |
| SUMIFS | `SUMIFS(求和范围, 条件范围1, 条件1, ...)` | 多条件求和 |
| ROUND | `ROUND(数值, 小数位数)` | 四舍五入 |
| ROUNDUP | `ROUNDUP(数值, 小数位数)` | 向上舍入 |
| ROUNDDOWN | `ROUNDDOWN(数值, 小数位数)` | 向下舍入 |
| SUMPRODUCT | `SUMPRODUCT(范围1, 范围2, ...)` | 对应项乘积之和 |
| MOD | `MOD(数值, 除数)` | 取余数 |

## 统计分析函数

| 函数 | 语法 | 说明 |
|------|------|------|
| AVERAGE | `AVERAGE(范围)` | 平均值 |
| AVERAGEIF | `AVERAGEIF(条件范围, 条件, 平均范围)` | 条件平均值 |
| COUNTIF | `COUNTIF(范围, 条件)` | 条件计数 |
| COUNTIFS | `COUNTIFS(条件范围1, 条件1, ...)` | 多条件计数 |
| MAX | `MAX(范围)` | 最大值 |
| MIN | `MIN(范围)` | 最小值 |
| RANK | `RANK(数值, 范围, 排序方式)` | 排名（0=降序, 1=升序） |

## 常用公式组合

### 多条件判断
```
=IF(AND(A1>=90,B1="完成"),"优秀",IF(AND(A1>=60,B1="完成"),"合格","待改进"))
```

### 模糊匹配查找
```
=XLOOKUP("*"&D1&"*",A:A,B:B,,2)
```

### 跨表求和
```
=SUMIF(Sheet2!A:A,A1,Sheet2!B:B)
```

### 动态区间查询
```
=VLOOKUP(A1,Sheet2!A:C,3,0)
→ XLOOKUP(A1,Sheet2!A:A,Sheet2!C:C)
```

### 日期区间计算
```
=NETWORKDAYS(A1,B1)  // 工作日
=DATEDIF(A1,B1,"Y") & "年" & DATEDIF(A1,B1,"YM") & "月"  // 年月
```

### 文本提取
```
=TRIM(MID(A1,FIND("【",A1)+1,FIND("】",A1)-FIND("【",A1)-1))
```
