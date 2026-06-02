# 配置文件詳解

## config_v2.json 結構

### 頂層字段

```json
{
  "version": "2.5",
  "schema_version": "1.4",
  "description": "甘特图生成器完整配置",
  "last_updated": "2026-05-12",
  "active_scheme": "blue_pro"
}
```

---

## 配色方案（schemes）

### 內置三種方案

```json
{
  "schemes": {
    "blue_pro": {
      "name": "專業藍調",
      "description": "參考VBA柔和藍色系",
      "colors": {
        "HEADER_BG": "1E3B4A",      // 標題背景色
        "HEADER_FG": "FFFFFF",      // 標題文字色
        "INFO_BG": "F0F4FA",        // 信息區背景
        "TASK_BG_END": "4A7A9C",    // 任務區結束色
        "MILESTONE_BG": "E5A155",   // 里程碑顏色
        "WEEKEND_BG": "E8E8E8",     // 周末背景色
        "TODAY_LINE": "D9534F",     // 今日線顏色
        "GRID": "C5D1DF",           // 網格線顏色
        "FONT_COLOR": "2C3E50"      // 字體顏色
      },
      "stage_colors": {
        "A": "2E86AB",  // A階段顏色
        "B": "81C784",  // B階段顏色
        "C": "FFD54F",  // C階段顏色
        "D": "E57373",  // D階段顏色
        "E": "9575CD",  // E階段顏色
        "F": "64B5F6",  // F階段顏色
        "G": "BFCF91"   // G階段顏色
      }
    }
  }
}
```

---

## 日曆配置（calendar）

```json
{
  "calendar": {
    "type": "ask_user",  // "ask_user" | "calendar_days" | "workdays"
    "options": ["calendar_days", "workdays"],
    "option_labels": {
      "calendar_days": "日曆天（包含周末与节假日）",
      "workdays": "工作天（週一至週五，排除假期）"
    },
    "work_week": [1,2,3,4,5],  // 工作日（1=週一, 5=週五）
    "holidays": ["2026-05-01", "2026-12-20", "2026-12-21"],
    "makeup_days": ["2026-05-09"]  // 補班日
  }
}
```

---

## 甘特圖布局（gantt_layout）

```json
{
  "gantt_layout": {
    "target_columns": 30,           // 目標列數
    "min_column_width": 4.0,        // 最小列寬
    "max_column_width": 7.0,        // 最大列寬
    "date_format": "MM/dd",         // 日期格式
    "month_format": "MM月",         // 月份格式
    "merge_months": true,           // 合併月份單元格
    "show_week_numbers": false,     // 顯示週數
    "auto_adjust_width": true       // 自動調整寬度
  }
}
```

---

## 數據驗證（data_validation）

```json
{
  "data_validation": {
    "min_duration": 1,              // 最小工期
    "max_duration": 90,             // 最大工期
    "default_duration": 1,          // 默認工期
    "default_predecessor_type": "FS",  // 默認前置類型
    "max_tasks": 200,               // 最大任務數
    "strict_validation": false      // 嚴格驗證模式
  }
}
```

---

## 輸出配置（output）

```json
{
  "output": {
    "file_pattern": "{project_name}_甘特图_{calendar_type}_{version}.xlsx",
    "overwrite": false,             // 覆蓋已存在文件
    "auto_open": false,             // 生成後自動打開
    "export_pdf": false,            // 導出 PDF
    "backup_enabled": true,         // 啟用備份
    "backup_count": 5               // 備份數量
  }
}
```

---

## 字段映射（fields）

定義 Excel 列標題到字段的自動映射：

```json
{
  "fields": {
    "name": {
      "patterns": ["施工內容", "項目", "任務名稱", "Description"],
      "type": "string",
      "required": true
    },
    "start": {
      "patterns": ["開始日期", "開始", "Start", "Sd"],
      "type": "date",
      "required": true
    },
    "duration": {
      "patterns": ["工期", "天數", "Duration", "dur"],
      "type": "integer",
      "default": 1
    }
  }
}
```

---

## 打印設置（print）

```json
{
  "print": {
    "fit_to_pages_wide": 1,         // 橫向適應頁數
    "fit_to_pages_tall": 0,         // 縱向不限
    "print_gridlines": true,        // 打印網格線
    "print_headings": true          // 打印標題
  }
}
```

---

## 自定義配色方案

創建自己的配色：

```json
{
  "schemes": {
    "my_custom": {
      "name": "自定義配色",
      "colors": {
        "HEADER_BG": "2C3E50",
        "HEADER_FG": "ECF0F1",
        "INFO_BG": "F5F5F5",
        "WEEKEND_BG": "E0E0E0",
        "TODAY_LINE": "E74C3C",
        "GRID": "BDC3C7",
        "FONT_COLOR": "2C3E50"
      },
      "stage_colors": {
        "A": "3498DB",
        "B": "2ECC71",
        "C": "F39C12",
        "D": "E74C3C"
      }
    }
  }
}
```

---

## 使用方法

1. 複製 `config_v2.json` 到工作目錄
2. 修改需要調整的字段
3. 運行腳本時自動加載

```bash
python scripts/gantt_chart_pro.py --input EDF.xlsx --config config_v2.json
```
