# 根因報告：PDF → Word 輸出落差分析

## 摘要

| 指標 | 自動生成 | 人工版本 | 差距 |
|------|----------|----------|------|
| Tables 數量 | 2 | 45+ | **95% 缺失** |
| Clauses 數量 | 675 (抽到但結構錯) | 675 | 結構不匹配 |
| 章節分表 | 無 (單一大表) | 有 (4/5/6/.../B) | 完全不同 |

## 根因分析

### 問題 1: 使用錯誤的模板

**現狀**
- 使用: `Generic-CB.docx` (2 個表格, 37KB)
- 正確: `AST-B-模板.docx` (45 個表格, 917KB)

**影響**
- 輸出的 Word 結構與人工版本完全不同
- 無法按章節分表 (4/5/6/7/8/9/10/B)
- 缺失附表 (M.3, M.4.2, T.7, X, 4.1.2...)

### 問題 2: Parser 的表格判斷邏輯過於嚴格

**現狀**
- `CBParser` 只靠 header 包含 "CLAUSE/VERDICT" 判斷
- MC-601.pdf 的表格 header 格式不完全匹配

**影響**
- 條款表抽取不穩定
- 舊版 CBParser 返回 0 clauses
- CBParserV2 改進後可抽到 675，但未按章節分群

### 問題 3: WordFiller 把複雜模板當單一表格處理

**現狀**
```python
# 清空所有數據行 (問題！)
while len(clause_table.rows) > 1:
    clause_table._tbl.remove(clause_table.rows[-1]._tr)
```

**影響**
- 破壞原有的多表格結構
- 無法保留章節分表
- 格式 (合併儲存格、樣式) 可能遺失

### 問題 4: 翻譯輸出包含 Assistant 話術

**現狀**
- 翻譯結果偶爾出現：「好的，請提供您要翻譯的英文原文…」
- 無 QA Gate 過濾錯誤輸出

**影響**
- 翻譯品質不穩定
- 需要人工檢查每個輸出

## 模板結構對比

### Generic-CB.docx (錯誤)
```
Tables: 2
  Table[0] header: Source (Overview)
  Table[1] header: Clause (單一大表)
```

### AST-B-模板.docx (正確)
```
Tables: 45
  Table[4]  header: 安全防護總攬表
  Table[5]  header: 能量源圖
  Table[6]  header: 4 (Chapter 4)
  Table[7]  header: 5 (Chapter 5)
  Table[8]  header: 6 (Chapter 6)
  Table[9]  header: 7 (Chapter 7)
  Table[10] header: 8 (Chapter 8)
  Table[11] header: 9 (Chapter 9)
  Table[12] header: 10 (Chapter 10)
  Table[13] header: B (Annex B)
  Table[36] header: M.3
  Table[37] header: M.4.2
  Table[44] header: 4.1.2
  ... (更多附表)
```

## 解決方案

### 階段 1: 基礎修復
1. 導入 AST-B-模板.docx 作為主要模板
2. 建立模板 registry 與 signature 系統
3. 修復 Parser 的表格判斷邏輯

### 階段 2: 結構化處理
4. Clauses 按章節分群 (clause_tables dict)
5. WordFiller 支援多表格回填
6. 附表抽取與回填

### 階段 3: 翻譯品質強化
7. Verdict deterministic mapping (P→符合)
8. QA Gate (過濾 Assistant 話術)
9. 去重 + 批次翻譯

### 階段 4: 驗證與部署
10. Coverage 驗證 (PDF vs Word)
11. 前端頁面整合
12. 部署優化

## 驗收標準

| 指標 | 目標值 |
|------|--------|
| 輸出 Tables 數量 | >= 40 |
| Clause 覆蓋率 | >= 95% |
| 章節分表正確性 | 100% |
| 翻譯 QA 通過率 | >= 98% |
| Verdict 映射正確率 | 100% |

---
生成時間: 2026-01-07
