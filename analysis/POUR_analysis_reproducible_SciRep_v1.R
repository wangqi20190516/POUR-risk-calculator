############################################################
# PUBLIC REPRODUCIBILITY VERSION
# Scientific Reports submission support
#
# Study:
# Dual-time-point prediction of early postoperative urinary
# retention after colpocleisis for advanced pelvic organ prolapse
#
# This public script is derived from the locked final analysis
# script used for the manuscript.
#
# PUBLIC-VERSION CHANGES ONLY:
# 1. Local Windows paths were replaced by project-relative paths.
# 2. Automatic package installation and interactive file selection
#    were removed.
# 3. Direct identifiers/non-model date fields are dropped, if present,
#    immediately after data import and are not required by the script.
# 4. glmnet was added to the dependency check because Section 16 uses it.
# 5. The run-information metadata was corrected to record the actual
#    5-fold penalized-regression sensitivity analysis.
# 6. Historical manuscript table/supplement numbering in comments and
#    output filenames was aligned with the current Scientific Reports
#    manuscript structure.
# 7. sessionInfo() is exported at the end for reproducibility.
#
# NO statistical algorithm, analysis variable, random seed,
# model-selection rule, bootstrap procedure, regression coefficient,
# or numerical calculation was intentionally changed.
#
# DATA:
# Place a de-identified analysis file named:
#     data/analysis_dataset.xlsx
# with worksheet:
#     Sheet1
# in the repository/project root before running this script.
#
# IMPORTANT:
# The repository should NOT contain the real patient-level dataset or
# the generated outputs/ directory. Those materials may contain
# sensitive clinical information.
############################################################



############################################################
# 1. 清空环境
############################################################

rm(list = ls())
graphics.off()
cat("\014")

options(stringsAsFactors = FALSE)
options(warn = 1)


############################################################
# 2. 检查并加载 R 包
############################################################

packages <- c(
  "dplyr",
  "tidyr",
  "readxl",
  "openxlsx",
  "mice",
  "ggplot2",
  "pROC",
  "car",
  "rms",
  "ResourceSelection",
  "glmnet"
)

missing_packages <- packages[
  !vapply(
    packages,
    requireNamespace,
    logical(1),
    quietly = TRUE
  )
]

if (length(missing_packages) > 0) {
  stop(
    paste0(
      "Missing required R packages: ",
      paste(missing_packages, collapse = ", "),
      ". Install them before running the analysis."
    )
  )
}

invisible(
  lapply(
    packages,
    library,
    character.only = TRUE
  )
)


############################################################
# 3. 设置路径并创建本次运行文件夹
############################################################

# ----------------------------------------------------------
# 3.1 De-identified analysis database (project-relative path)
# ----------------------------------------------------------

project_root <- normalizePath(
  ".",
  winslash = "/",
  mustWork = TRUE
)

data_file <- file.path(
  project_root,
  "data",
  "analysis_dataset.xlsx"
)

data_sheet <- "Sheet1"


# ----------------------------------------------------------
# 3.2 Output directory (project-relative path)
# ----------------------------------------------------------

output_base <- file.path(
  project_root,
  "outputs"
)

if (!file.exists(data_file)) {
  stop(
    paste0(
      "Cannot find the de-identified analysis dataset at: ",
      data_file,
      "\nPlace analysis_dataset.xlsx in the data/ directory."
    )
  )
}

if (!dir.exists(output_base)) {
  dir.create(
    output_base,
    recursive = TRUE
  )
}


# ----------------------------------------------------------
# 3.3 创建本次独立运行文件夹
# ----------------------------------------------------------

run_timestamp <- format(
  Sys.time(),
  "%Y%m%d_%H%M%S"
)

output_dir <- file.path(
  output_base,
  paste0(
    "POUR_traditional_",
    run_timestamp
  )
)


# ----------------------------------------------------------
# 3.4 创建本次运行的各类结果文件夹
# ----------------------------------------------------------

data_dir <- file.path(
  output_dir,
  "01_data"
)

table_dir <- file.path(
  output_dir,
  "02_tables"
)

figure_dir <- file.path(
  output_dir,
  "03_figures"
)

model_dir <- file.path(
  output_dir,
  "04_models"
)

log_dir <- file.path(
  output_dir,
  "05_logs"
)


dir.create(
  output_dir,
  recursive = TRUE
)

dir.create(
  data_dir,
  showWarnings = FALSE
)

dir.create(
  table_dir,
  showWarnings = FALSE
)

dir.create(
  figure_dir,
  showWarnings = FALSE
)

dir.create(
  model_dir,
  showWarnings = FALSE
)

dir.create(
  log_dir,
  showWarnings = FALSE
)


# ----------------------------------------------------------
# 3.5 记录本次运行基本信息
# ----------------------------------------------------------

run_info <- data.frame(
  run_timestamp = run_timestamp,
  input_file = file.path(
    "data",
    "analysis_dataset.xlsx"
  ),
  data_sheet = data_sheet,
  output_dir = file.path(
    "outputs",
    basename(output_dir)
  ),
  univariable_cutoff = 0.10,
  multivariable_cutoff = 0.05,
  cross_validation_folds = 5,
  stringsAsFactors = FALSE
)


openxlsx::write.xlsx(
  run_info,
  file.path(
    output_dir,
    "run_information.xlsx"
  ),
  overwrite = TRUE
)


cat("\n本次运行输入数据库：\n")
cat(data_file, "\n")

cat("\n本次运行输出文件夹：\n")
cat(output_dir, "\n\n")


############################################################
# 4. 定义变量
############################################################

outcome_var <- "pour"
id_var <- "patient_id"

# 以下变量仅用于病例识别或时间标记，不进入任何预测模型
non_model_vars <- c(
  "patient_id"
)

# Direct identifiers and non-model dates are not required for analysis.
# If present in an authorized source file, remove them immediately after import.
privacy_columns_to_drop <- c(
  "patient_name",
  "surgery_date"
)


# ----------------------------------------------------------
# 4.1 术前时间点允许考虑的全部预测变量
#
# 所有变量均进入描述性/组间单因素比较。
#
# 是否进入后续多因素模型，由 observed-data
# between-group comparison P < 0.10 决定。
#
# 注意：
# 术前模型不得纳入任何术中或拔管前变量。
# ----------------------------------------------------------

preoperative_screening_vars <- c(
  "age",
  "years_since_menopause",
  "gravidity",
  "parity",
  "height_cm",
  "weight_kg",
  "bmi",
  "prolapse_duration",
  "prior_hysterectomy",
  "hypertension",
  "diabetes",
  "insulin_use",
  "coronary_heart_disease",
  "uterine_prolapse_stage",
  "anterior_wall_stage",
  "posterior_wall_stage",
  "popq_aa",
  "popq_ba",
  "popq_c",
  "popq_ap",
  "popq_bp",
  "voiding_difficulty",
  "bowel_difficulty",
  "frequency",
  "urgency",
  "incomplete_emptying",
  "preop_pvr",
  "detrusor_thickening",
  "resting_posterior_angle",
  "resting_urethral_inclination_angle",
  "urethral_funneling",
  "valsalva_posterior_angle",
  "valsalva_urethral_inclination_angle",
  "urethral_rotation_angle",
  "bladder_neck_descent",
  "cystocele_type"
)


# ----------------------------------------------------------
# 4.2 拔管前时间点允许考虑的全部预测变量
#
# 拔管前模型：
# 全部术前变量
# +
# 在首次拔管前已经获得的围手术期信息
# ----------------------------------------------------------

precatheter_additional_vars <- c(
  "colpocleisis_type",
  "concomitant_hysterectomy",
  "operative_time_min",
  "blood_loss_ml",
  "time_to_first_catheter_removal_pod"
)

precatheter_screening_vars <- c(
  preoperative_screening_vars,
  precatheter_additional_vars
)


# Table 1 与单因素 Logistic 使用相同的时间点变量池
preoperative_table_vars <- preoperative_screening_vars

precatheter_table_vars <- precatheter_screening_vars


# ----------------------------------------------------------
# 4.3 连续变量
#
# 这些变量在 Table 1 中根据分布情况决定：
# 正态分布    → Mean ± SD + t test
# 非正态分布  → Median (IQR) + Mann-Whitney U test
#
# 在 Logistic 模型中原则上保留连续形式，
# 不预先人为二分类。
# ----------------------------------------------------------

continuous_vars <- c(
  "age",
  "years_since_menopause",
  "gravidity",
  "parity",
  "height_cm",
  "weight_kg",
  "bmi",
  "prolapse_duration",
  "popq_aa",
  "popq_ba",
  "popq_c",
  "popq_ap",
  "popq_bp",
  "preop_pvr",
  "resting_posterior_angle",
  "resting_urethral_inclination_angle",
  "valsalva_posterior_angle",
  "valsalva_urethral_inclination_angle",
  "urethral_rotation_angle",
  "bladder_neck_descent",
  "operative_time_min",
  "blood_loss_ml",
  "time_to_first_catheter_removal_pod"
)


# ----------------------------------------------------------
# 4.4 二分类变量
# ----------------------------------------------------------

# 普通 Yes / No 二分类变量
# 原始数据库编码：
# 0 = No
# 1 = Yes

binary_yes_no_vars <- c(
  "prior_hysterectomy",
  "hypertension",
  "diabetes",
  "insulin_use",
  "coronary_heart_disease",
  "voiding_difficulty",
  "bowel_difficulty",
  "frequency",
  "urgency",
  "incomplete_emptying",
  "detrusor_thickening",
  "urethral_funneling",
  "concomitant_hysterectomy"
)


# 特殊二分类变量
# 也是0/1编码，但含义不是No/Yes

binary_special_vars <- c(
  "colpocleisis_type"
)


# 所有二分类变量
binary_vars <- c(
  binary_yes_no_vars,
  binary_special_vars
)

# ----------------------------------------------------------
# 4.5 多分类变量
#
# 主分析按名义分类变量（nominal factor）处理，
# 不按普通连续变量处理，也不使用 ordered factor。
#
# 这样不预设：
# “从II期到III期”与“III期到IV期”的效应必须相同。
# ----------------------------------------------------------

nominal_factor_vars <- c(
  "uterine_prolapse_stage",
  "anterior_wall_stage",
  "posterior_wall_stage",
  "cystocele_type"
)


# 所有分类变量
categorical_vars <- c(
  binary_vars,
  nominal_factor_vars
)


# ----------------------------------------------------------
# 4.6 两个预测时间点各自涉及的连续/分类变量
# ----------------------------------------------------------

preoperative_continuous_vars <- intersect(
  continuous_vars,
  preoperative_screening_vars
)

preoperative_categorical_vars <- intersect(
  categorical_vars,
  preoperative_screening_vars
)

precatheter_continuous_vars <- intersect(
  continuous_vars,
  precatheter_screening_vars
)

precatheter_categorical_vars <- intersect(
  categorical_vars,
  precatheter_screening_vars
)


# ----------------------------------------------------------
# 4.7 可能存在明显同源或高度相关信息的变量组
#
# 重要：
# 这里只是定义“需要比较”的变量组，
# 不在此阶段自动删除任何变量，
# 也不预设哪一个变量必须优先保留。
#
# 如果同一组中两个或以上变量均达到
# 单因素 Logistic P < 0.10，
# 后续分别建立候选多因素模型，
# 比较 AUC、AIC、BIC 等模型表现，
# 再结合 VIF/GVIF 检查共线性，
# 确定最终采用哪一种变量表示。
# ----------------------------------------------------------

related_variable_groups <- list(
  
  body_size = c(
    "height_cm",
    "weight_kg",
    "bmi"
  ),
  
  anterior_prolapse = c(
    "anterior_wall_stage",
    "popq_aa",
    "popq_ba"
  ),
  
  apical_prolapse = c(
    "uterine_prolapse_stage",
    "popq_c"
  ),
  
  posterior_prolapse = c(
    "posterior_wall_stage",
    "popq_ap",
    "popq_bp"
  ),
  
  urethral_mobility = c(
    "resting_urethral_inclination_angle",
    "valsalva_urethral_inclination_angle",
    "urethral_rotation_angle"
  )
)


# ----------------------------------------------------------
# 4.8 本研究当前不人为合并的变量
#
# 下列症状均保持独立：
# voiding_difficulty
# incomplete_emptying
# frequency
# urgency
#
# 不再生成：
# preop_emptying_symptom
# storage_symptom
#
# preop_pvr 保持连续数值，
# 不预先二分类为 >=100 mL 或 >=200 mL。
# ----------------------------------------------------------



############################################################
# 5. 定义通用辅助函数
############################################################


# ----------------------------------------------------------
# 5.1 保存结果表格
#
# 每张结果表同时保存为：
# 1. Excel (.xlsx)
# 2. CSV (.csv)
# ----------------------------------------------------------

save_table <- function(
    data,
    file_name,
    folder = table_dir
) {
  
  openxlsx::write.xlsx(
    data,
    file.path(
      folder,
      paste0(file_name, ".xlsx")
    ),
    overwrite = TRUE
  )
  
  write.csv(
    data,
    file.path(
      folder,
      paste0(file_name, ".csv")
    ),
    row.names = FALSE,
    na = "",
    fileEncoding = "UTF-8"
  )
}


# ----------------------------------------------------------
# 5.2 保存 ggplot 图
#
# 同时保存高分辨率 JPEG 和 PDF。
# ----------------------------------------------------------

save_plot <- function(
    plot_object,
    file_name,
    width = 8,
    height = 6
) {
  
  ggplot2::ggsave(
    filename = file.path(
      figure_dir,
      paste0(file_name, ".jpeg")
    ),
    plot = plot_object,
    width = width,
    height = height,
    dpi = 600,
    bg = "white"
  )
  
  ggplot2::ggsave(
    filename = file.path(
      figure_dir,
      paste0(file_name, ".pdf")
    ),
    plot = plot_object,
    width = width,
    height = height,
    bg = "white"
  )
}


# ----------------------------------------------------------
# 5.3 P值格式
#
# P < 0.001 显示为 "<0.001"
# 其余保留三位小数。
# ----------------------------------------------------------

format_p_value <- function(p) {
  
  if (is.na(p)) {
    return("")
  }
  
  if (p < 0.001) {
    return("<0.001")
  }
  
  sprintf("%.3f", p)
}


# ----------------------------------------------------------
# 5.4 0/1变量转换为二分类因子
#
# 默认：
# 0 = No
# 1 = Yes
#
# 对特殊变量可以自行指定标签。
# ----------------------------------------------------------

make_binary_factor <- function(
    x,
    labels = c("No", "Yes")
) {
  
  # 只允许 0、1 和缺失值
  invalid_values <- unique(
    x[
      !is.na(x) &
        !(x %in% c(0, 1))
    ]
  )
  
  if (length(invalid_values) > 0) {
    
    stop(
      paste0(
        "发现非0/1编码：",
        paste(
          invalid_values,
          collapse = ", "
        )
      )
    )
  }
  
  factor(
    x,
    levels = c(0, 1),
    labels = labels
  )
}

############################################################
# 6. 读取正式数据库并完成变量编码
############################################################


# ----------------------------------------------------------
# 6.1 读取 Excel
# ----------------------------------------------------------

raw_data <- readxl::read_xlsx(
  data_file,
  sheet = data_sheet
)

analysis_data <- as.data.frame(
  raw_data,
  check.names = FALSE
)


privacy_columns_present <- intersect(
  privacy_columns_to_drop,
  names(analysis_data)
)

if (length(privacy_columns_present) > 0) {
  analysis_data <- analysis_data[
    ,
    setdiff(
      names(analysis_data),
      privacy_columns_present
    ),
    drop = FALSE
  ]
}


# ----------------------------------------------------------
# 6.2 检查数据库字段
#
# 这里检查的是Excel原始字段。
# 三个 *_stage_model 是后面在R中生成的，
# 因此不属于Excel原始字段。
# ----------------------------------------------------------

expected_input_vars <- unique(
  c(
    non_model_vars,
    outcome_var,
    preoperative_table_vars,
    precatheter_additional_vars
  )
)

missing_columns <- setdiff(
  expected_input_vars,
  names(analysis_data)
)

unexpected_columns <- setdiff(
  names(analysis_data),
  expected_input_vars
)

if (length(missing_columns) > 0) {
  
  stop(
    paste0(
      "Excel 中缺少以下必要字段：",
      paste(
        missing_columns,
        collapse = ", "
      )
    )
  )
}

if (length(unexpected_columns) > 0) {
  
  stop(
    paste0(
      "Excel 中发现预期之外的字段：",
      paste(
        unexpected_columns,
        collapse = ", "
      )
    )
  )
}


# ----------------------------------------------------------
# 6.3 核对锁库数据库基本结构
#
# 当前正式锁库版本：
# 567例，50例POUR
# ----------------------------------------------------------

expected_sample_n <- 567
expected_pour_events <- 50

if (nrow(analysis_data) != expected_sample_n) {
  
  stop(
    paste0(
      "样本量与锁库数据库不一致：当前为 ",
      nrow(analysis_data),
      "，预期为 ",
      expected_sample_n
    )
  )
}

if (
  any(is.na(analysis_data[[id_var]])) ||
  anyDuplicated(analysis_data[[id_var]]) > 0
) {
  
  stop(
    "patient_id 存在缺失或重复。"
  )
}

if (
  !identical(
    sort(
      as.integer(
        analysis_data[[id_var]]
      )
    ),
    1:567
  )
) {
  
  stop(
    "patient_id 应完整覆盖 1–567。"
  )
}


# ----------------------------------------------------------
# 6.4 严格检查结局变量
#
# pour保持数值0/1，
# 不另外建立pour_binary。
# ----------------------------------------------------------

if (!is.numeric(analysis_data[[outcome_var]])) {
  
  stop(
    "pour 应为数值型0/1，请检查Excel数据库。"
  )
}

invalid_outcome <- unique(
  analysis_data[[outcome_var]][
    !is.na(analysis_data[[outcome_var]]) &
      !(analysis_data[[outcome_var]] %in% c(0, 1))
  ]
)

if (length(invalid_outcome) > 0) {
  
  stop(
    paste0(
      "pour 中存在非0/1编码：",
      paste(
        invalid_outcome,
        collapse = ", "
      )
    )
  )
}

if (any(is.na(analysis_data[[outcome_var]]))) {
  
  stop(
    "pour 结局存在缺失值。"
  )
}

analysis_data[[outcome_var]] <- as.integer(
  analysis_data[[outcome_var]]
)

current_pour_events <- sum(
  analysis_data[[outcome_var]] == 1
)

if (current_pour_events != expected_pour_events) {
  
  stop(
    paste0(
      "POUR事件数与锁库数据库不一致：当前为 ",
      current_pour_events,
      "，预期为 ",
      expected_pour_events
    )
  )
}


# ----------------------------------------------------------
# 6.5 严格检查连续变量
#
# 正式锁库以后不再使用as.numeric()自动纠错。
# 如果某个连续变量被读取为字符型，
# 直接停止并提示检查数据库。
# ----------------------------------------------------------

non_numeric_continuous_vars <- continuous_vars[
  !vapply(
    analysis_data[
      continuous_vars
    ],
    is.numeric,
    logical(1)
  )
]

if (length(non_numeric_continuous_vars) > 0) {
  
  stop(
    paste0(
      "以下连续变量未被R识别为数值型：",
      paste(
        non_numeric_continuous_vars,
        collapse = ", "
      )
    )
  )
}


# ----------------------------------------------------------
# 6.6 将普通0/1二分类变量转为 No / Yes
# ----------------------------------------------------------

for (v in binary_yes_no_vars) {
  
  analysis_data[[v]] <- make_binary_factor(
    analysis_data[[v]],
    labels = c(
      "No",
      "Yes"
    )
  )
}


# ----------------------------------------------------------
# 6.7 单独编码封闭术方式
#
# 0 = LeFort
# 1 = Modified colpocleisis
# ----------------------------------------------------------

analysis_data$colpocleisis_type <- make_binary_factor(
  analysis_data$colpocleisis_type,
  labels = c(
    "LeFort",
    "Modified colpocleisis"
  )
)


# ----------------------------------------------------------
# 6.8 检查三个原始POP分期
#
# 原数据库只能为：
# 0、1、2、3、4
# ----------------------------------------------------------

original_stage_vars <- c(
  "uterine_prolapse_stage",
  "anterior_wall_stage",
  "posterior_wall_stage"
)

for (v in original_stage_vars) {
  
  if (!is.numeric(analysis_data[[v]])) {
    
    stop(
      paste0(
        v,
        " 应为数值型0–4。"
      )
    )
  }
  
  invalid_stage <- unique(
    analysis_data[[v]][
      !is.na(analysis_data[[v]]) &
        !(analysis_data[[v]] %in% 0:4)
    ]
  )
  
  if (length(invalid_stage) > 0) {
    
    stop(
      paste0(
        v,
        " 中存在0–4之外的分期：",
        paste(
          invalid_stage,
          collapse = ", "
        )
      )
    )
  }
}


# ----------------------------------------------------------
# 6.9 生成建模用POP分期
#
# 原始数据库不修改。
#
# Table 1：
# 0、I、II、III、IV
#
# Logistic模型：
# 0-I、II、III、IV
# ----------------------------------------------------------

make_stage_model <- function(x) {
  
  stage_model <- ifelse(
    x %in% c(0, 1),
    "0-I",
    ifelse(
      x == 2,
      "II",
      ifelse(
        x == 3,
        "III",
        ifelse(
          x == 4,
          "IV",
          NA_character_
        )
      )
    )
  )
  
  factor(
    stage_model,
    levels = c(
      "0-I",
      "II",
      "III",
      "IV"
    )
  )
}


analysis_data$uterine_prolapse_stage_model <- make_stage_model(
  analysis_data$uterine_prolapse_stage
)

analysis_data$anterior_wall_stage_model <- make_stage_model(
  analysis_data$anterior_wall_stage
)

analysis_data$posterior_wall_stage_model <- make_stage_model(
  analysis_data$posterior_wall_stage
)


# ----------------------------------------------------------
# 6.10 原始分期转换为Table 1展示用factor
#
# 原始信息仍完整保留：
# 0、I、II、III、IV
# ----------------------------------------------------------

stage_display_levels <- c(
  "0",
  "I",
  "II",
  "III",
  "IV"
)

for (v in original_stage_vars) {
  
  analysis_data[[v]] <- factor(
    analysis_data[[v]],
    levels = 0:4,
    labels = stage_display_levels
  )
}


# ----------------------------------------------------------
# 6.11 膀胱膨出类型
#
# 当前数据库允许：
# 0、2、3、缺失
#
# 作为名义分类变量处理，
# 不视为连续变量。
# ----------------------------------------------------------

if (!is.numeric(analysis_data$cystocele_type)) {
  
  stop(
    "cystocele_type 应为数值编码0、2、3或缺失。"
  )
}

invalid_cystocele_type <- unique(
  analysis_data$cystocele_type[
    !is.na(analysis_data$cystocele_type) &
      !(analysis_data$cystocele_type %in% c(0, 2, 3))
  ]
)

if (length(invalid_cystocele_type) > 0) {
  
  stop(
    paste0(
      "cystocele_type 中存在异常编码：",
      paste(
        invalid_cystocele_type,
        collapse = ", "
      )
    )
  )
}

analysis_data$cystocele_type <- factor(
  analysis_data$cystocele_type,
  levels = c(
    0,
    2,
    3
  ),
  labels = c(
    "0",
    "II",
    "III"
  )
)


# ----------------------------------------------------------
# 6.12 创建Table 1显示用结局分组
#
# 不用于Logistic建模。
#
# Table 1要求：
# Overall | POUR | Non-POUR
# ----------------------------------------------------------

analysis_data$pour_group <- factor(
  analysis_data$pour,
  levels = c(
    1,
    0
  ),
  labels = c(
    "POUR",
    "Non-POUR"
  )
)


# ----------------------------------------------------------
# 6.13 最终核查
# ----------------------------------------------------------

cat("\n数据库读取和编码完成。\n\n")

cat(
  "总样本量：",
  nrow(analysis_data),
  "\n"
)

cat(
  "POUR事件数：",
  sum(analysis_data$pour == 1),
  "\n"
)

cat(
  "Non-POUR例数：",
  sum(analysis_data$pour == 0),
  "\n"
)

cat(
  "POUR发生率：",
  sprintf(
    "%.2f%%",
    100 *
      mean(
        analysis_data$pour
      )
  ),
  "\n\n"
)

cat("建模用POP分期分布：\n")

print(
  table(
    analysis_data$uterine_prolapse_stage_model,
    useNA = "ifany"
  )
)

print(
  table(
    analysis_data$anterior_wall_stage_model,
    useNA = "ifany"
  )
)

print(
  table(
    analysis_data$posterior_wall_stage_model,
    useNA = "ifany"
  )
)


# ----------------------------------------------------------
# 6.14 保存编码后的R数据对象
# ----------------------------------------------------------

saveRDS(
  analysis_data,
  file.path(
    data_dir,
    "analysis_data_encoded.rds"
  )
)


############################################################
# 7. 缺失值报告
############################################################


# ----------------------------------------------------------
# 7.1 全数据库缺失情况
# ----------------------------------------------------------

missing_report <- data.frame(
  variable = names(analysis_data),
  variable_class = sapply(
    analysis_data,
    function(x) {
      paste(
        class(x),
        collapse = "/"
      )
    }
  ),
  missing_n = sapply(
    analysis_data,
    function(x) {
      sum(is.na(x))
    }
  ),
  stringsAsFactors = FALSE
)

missing_report$missing_percent <- round(
  missing_report$missing_n /
    nrow(analysis_data) *
    100,
  2
)

missing_report <- missing_report %>%
  arrange(
    desc(missing_percent),
    variable
  )


# ----------------------------------------------------------
# 7.2 术前时间点缺失情况
#
# 使用Table 1中的原始术前变量。
# 三个 *_stage_model 为R中派生变量，
# 不单独作为原始缺失数据报告对象。
# ----------------------------------------------------------

preoperative_missing_report <- missing_report %>%
  filter(
    variable %in%
      c(
        outcome_var,
        preoperative_table_vars
      )
  )


# ----------------------------------------------------------
# 7.3 拔管前时间点缺失情况
# ----------------------------------------------------------

precatheter_missing_report <- missing_report %>%
  filter(
    variable %in%
      c(
        outcome_var,
        precatheter_table_vars
      )
  )


# ----------------------------------------------------------
# 7.4 保存缺失值报告
# ----------------------------------------------------------

openxlsx::write.xlsx(
  list(
    All_variables = missing_report,
    Preoperative = preoperative_missing_report,
    Precatheter = precatheter_missing_report
  ),
  file.path(
    data_dir,
    "Missing_value_report.xlsx"
  ),
  overwrite = TRUE
)

write.csv(
  missing_report,
  file.path(
    data_dir,
    "Missing_value_report.csv"
  ),
  row.names = FALSE,
  na = "",
  fileEncoding = "UTF-8"
)


# ----------------------------------------------------------
# 7.5 安全检查
# ----------------------------------------------------------

fully_missing_vars <- missing_report$variable[
  missing_report$missing_n ==
    nrow(analysis_data)
]

if (length(fully_missing_vars) > 0) {
  
  stop(
    paste0(
      "发现整列完全缺失的变量：",
      paste(
        fully_missing_vars,
        collapse = ", "
      )
    )
  )
}


# ----------------------------------------------------------
# 7.6 控制台显示真正存在缺失的变量
# ----------------------------------------------------------

cat("\n存在缺失值的变量：\n")

print(
  missing_report %>%
    filter(
      missing_n > 0
    )
)

cat(
  "\n存在缺失值的变量数：",
  sum(
    missing_report$missing_n > 0
  ),
  "\n\n"
)



############################################################
# 8. MICE 多重插补
#
# 分别建立：
# 1. 术前时间点 MICE 对象
# 2. 首次拔管前时间点 MICE 对象
#
# 正式主分析：
# m = 20 个插补数据集
# maxit = 20 次迭代
############################################################


# ----------------------------------------------------------
# 8.1 MICE基本设置
# ----------------------------------------------------------

mice_m <- 20
mice_maxit <- 20

mice_seed_preoperative <- 20260817
mice_seed_precatheter <- 20260818


# ----------------------------------------------------------
# 8.2 明确需要插补的变量及方法
#
# 当前共有10个变量存在缺失：
#
# 连续变量 → pmm
# 二分类变量 → logreg
# 多分类变量 → polyreg
# ----------------------------------------------------------

mice_numeric_missing_vars <- c(
  "preop_pvr",
  "resting_posterior_angle",
  "resting_urethral_inclination_angle",
  "valsalva_posterior_angle",
  "valsalva_urethral_inclination_angle",
  "urethral_rotation_angle",
  "bladder_neck_descent"
)

mice_binary_missing_vars <- c(
  "detrusor_thickening",
  "urethral_funneling"
)

mice_nominal_missing_vars <- c(
  "cystocele_type"
)

mice_missing_vars <- c(
  mice_numeric_missing_vars,
  mice_binary_missing_vars,
  mice_nominal_missing_vars
)


# ----------------------------------------------------------
# 8.3 确认上述10个变量确实存在缺失
# ----------------------------------------------------------

mice_missing_check <- data.frame(
  variable = mice_missing_vars,
  missing_n = sapply(
    analysis_data[mice_missing_vars],
    function(x) {
      sum(is.na(x))
    }
  ),
  stringsAsFactors = FALSE
)

print(
  mice_missing_check
)

if (
  any(
    mice_missing_check$missing_n == 0
  )
) {
  
  warning(
    "预设MICE缺失变量中存在当前没有缺失的变量，请检查。"
  )
}


# ----------------------------------------------------------
# 8.4 建立术前MICE数据
#
# pour：
# - 不插补
# - 可以作为其他缺失变量的预测因子
#
# 不包含：
# patient_id
# 以及任何术中/拔管前变量
# ----------------------------------------------------------

preoperative_mice_vars <- unique(
  c(
    outcome_var,
    preoperative_screening_vars
  )
)

preoperative_mice_data <- analysis_data[
  preoperative_mice_vars
]


# ----------------------------------------------------------
# 8.5 设置术前MICE插补方法
# ----------------------------------------------------------

preoperative_method <- mice::make.method(
  preoperative_mice_data
)

# 先将所有变量设为不插补
preoperative_method[] <- ""

# 连续缺失变量
for (v in mice_numeric_missing_vars) {
  
  if (v %in% names(preoperative_mice_data)) {
    preoperative_method[v] <- "pmm"
  }
}

# 二分类缺失变量
for (v in mice_binary_missing_vars) {
  
  if (v %in% names(preoperative_mice_data)) {
    preoperative_method[v] <- "logreg"
  }
}

# 多分类缺失变量
for (v in mice_nominal_missing_vars) {
  
  if (v %in% names(preoperative_mice_data)) {
    preoperative_method[v] <- "polyreg"
  }
}

# POUR结局绝不插补
preoperative_method[outcome_var] <- ""


# ----------------------------------------------------------
# 8.6 设置术前MICE predictor matrix
# ----------------------------------------------------------

preoperative_predictor_matrix <- mice::make.predictorMatrix(
  preoperative_mice_data
)

# 一个变量不能预测自己
diag(
  preoperative_predictor_matrix
) <- 0

# 不需要插补的变量，其“目标变量行”设为0
non_imputed_preoperative_vars <- names(
  preoperative_method[
    preoperative_method == ""
  ]
)

preoperative_predictor_matrix[
  non_imputed_preoperative_vars,
] <- 0

# POUR本身不被插补
preoperative_predictor_matrix[
  outcome_var,
] <- 0

# 但POUR允许作为10个缺失变量的辅助预测因子
preoperative_imputed_targets <- names(
  preoperative_method[
    preoperative_method != ""
  ]
)

preoperative_predictor_matrix[
  preoperative_imputed_targets,
  outcome_var
] <- 1


# ----------------------------------------------------------
# 8.7 运行术前MICE
# ----------------------------------------------------------

cat(
  "\n开始术前时间点 MICE 多重插补...\n"
)

set.seed(
  mice_seed_preoperative
)

imp_preoperative <- mice::mice(
  data = preoperative_mice_data,
  m = mice_m,
  maxit = mice_maxit,
  method = preoperative_method,
  predictorMatrix = preoperative_predictor_matrix,
  seed = mice_seed_preoperative,
  printFlag = TRUE
)

cat(
  "\n术前时间点 MICE 完成。\n"
)


# ----------------------------------------------------------
# 8.8 建立拔管前MICE数据
#
# 可以使用：
# 全部术前变量
# +
# 5个拔管前已经获得的变量
# ----------------------------------------------------------

precatheter_mice_vars <- unique(
  c(
    outcome_var,
    precatheter_screening_vars
  )
)

precatheter_mice_data <- analysis_data[
  precatheter_mice_vars
]


# ----------------------------------------------------------
# 8.9 设置拔管前MICE插补方法
# ----------------------------------------------------------

precatheter_method <- mice::make.method(
  precatheter_mice_data
)

precatheter_method[] <- ""

# 连续缺失变量
for (v in mice_numeric_missing_vars) {
  
  if (v %in% names(precatheter_mice_data)) {
    precatheter_method[v] <- "pmm"
  }
}

# 二分类缺失变量
for (v in mice_binary_missing_vars) {
  
  if (v %in% names(precatheter_mice_data)) {
    precatheter_method[v] <- "logreg"
  }
}

# 多分类缺失变量
for (v in mice_nominal_missing_vars) {
  
  if (v %in% names(precatheter_mice_data)) {
    precatheter_method[v] <- "polyreg"
  }
}

# POUR绝不插补
precatheter_method[outcome_var] <- ""


# ----------------------------------------------------------
# 8.10 设置拔管前MICE predictor matrix
# ----------------------------------------------------------

precatheter_predictor_matrix <- mice::make.predictorMatrix(
  precatheter_mice_data
)

diag(
  precatheter_predictor_matrix
) <- 0

non_imputed_precatheter_vars <- names(
  precatheter_method[
    precatheter_method == ""
  ]
)

precatheter_predictor_matrix[
  non_imputed_precatheter_vars,
] <- 0

precatheter_predictor_matrix[
  outcome_var,
] <- 0

precatheter_imputed_targets <- names(
  precatheter_method[
    precatheter_method != ""
  ]
)

precatheter_predictor_matrix[
  precatheter_imputed_targets,
  outcome_var
] <- 1


# ----------------------------------------------------------
# 8.11 运行拔管前MICE
# ----------------------------------------------------------

cat(
  "\n开始拔管前时间点 MICE 多重插补...\n"
)

set.seed(
  mice_seed_precatheter
)

imp_precatheter <- mice::mice(
  data = precatheter_mice_data,
  m = mice_m,
  maxit = mice_maxit,
  method = precatheter_method,
  predictorMatrix = precatheter_predictor_matrix,
  seed = mice_seed_precatheter,
  printFlag = TRUE
)

cat(
  "\n拔管前时间点 MICE 完成。\n"
)


# ----------------------------------------------------------
# 8.12 保存两个MICE对象
#
# 后续所有正式回归分析均从这两个对象继续，
# 不再建立唯一的completed_data。
# ----------------------------------------------------------

saveRDS(
  imp_preoperative,
  file.path(
    data_dir,
    "MICE_preoperative_m20_maxit20.rds"
  )
)

saveRDS(
  imp_precatheter,
  file.path(
    data_dir,
    "MICE_precatheter_m20_maxit20.rds"
  )
)


# ----------------------------------------------------------
# 8.13 输出MICE方法设置
# ----------------------------------------------------------

preoperative_method_table <- data.frame(
  variable = names(
    preoperative_method
  ),
  missing_n = sapply(
    preoperative_mice_data,
    function(x) {
      sum(is.na(x))
    }
  ),
  method = unname(
    preoperative_method
  ),
  stringsAsFactors = FALSE
)

precatheter_method_table <- data.frame(
  variable = names(
    precatheter_method
  ),
  missing_n = sapply(
    precatheter_mice_data,
    function(x) {
      sum(is.na(x))
    }
  ),
  method = unname(
    precatheter_method
  ),
  stringsAsFactors = FALSE
)


# ----------------------------------------------------------
# 8.14 转换predictor matrix为可查看的表格
# ----------------------------------------------------------

preoperative_predictor_table <- data.frame(
  target_variable = rownames(
    preoperative_predictor_matrix
  ),
  preoperative_predictor_matrix,
  row.names = NULL,
  check.names = FALSE
)

precatheter_predictor_table <- data.frame(
  target_variable = rownames(
    precatheter_predictor_matrix
  ),
  precatheter_predictor_matrix,
  row.names = NULL,
  check.names = FALSE
)


# ----------------------------------------------------------
# 8.15 保存MICE设置
# ----------------------------------------------------------

openxlsx::write.xlsx(
  list(
    Preop_method =
      preoperative_method_table,
    Preop_predictor_matrix =
      preoperative_predictor_table,
    Precatheter_method =
      precatheter_method_table,
    Precatheter_predictor_matrix =
      precatheter_predictor_table
  ),
  file.path(
    data_dir,
    "MICE_settings.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 8.16 保存MICE运行警告/事件
# ----------------------------------------------------------

preoperative_logged_events <- (
  imp_preoperative$loggedEvents
)

precatheter_logged_events <- (
  imp_precatheter$loggedEvents
)

if (
  is.null(
    preoperative_logged_events
  )
) {
  
  preoperative_logged_events <- data.frame(
    message =
      "No logged events"
  )
}

if (
  is.null(
    precatheter_logged_events
  )
) {
  
  precatheter_logged_events <- data.frame(
    message =
      "No logged events"
  )
}

openxlsx::write.xlsx(
  list(
    Preoperative =
      preoperative_logged_events,
    Precatheter =
      precatheter_logged_events
  ),
  file.path(
    log_dir,
    "MICE_logged_events.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 8.17 检查20个插补数据集是否仍存在缺失
# ----------------------------------------------------------

for (i in 1:mice_m) {
  
  completed_preoperative_i <- mice::complete(
    imp_preoperative,
    action = i
  )
  
  if (
    anyNA(
      completed_preoperative_i
    )
  ) {
    
    stop(
      paste0(
        "术前MICE第",
        i,
        "个插补数据集仍存在缺失值。"
      )
    )
  }
}

for (i in 1:mice_m) {
  
  completed_precatheter_i <- mice::complete(
    imp_precatheter,
    action = i
  )
  
  if (
    anyNA(
      completed_precatheter_i
    )
  ) {
    
    stop(
      paste0(
        "拔管前MICE第",
        i,
        "个插补数据集仍存在缺失值。"
      )
    )
  }
}


# ----------------------------------------------------------
# 8.18 保存MICE收敛诊断图
# ----------------------------------------------------------

pdf(
  file.path(
    log_dir,
    "MICE_convergence_preoperative.pdf"
  ),
  width = 11,
  height = 8
)

plot(
  imp_preoperative
)

dev.off()


pdf(
  file.path(
    log_dir,
    "MICE_convergence_precatheter.pdf"
  ),
  width = 11,
  height = 8
)

plot(
  imp_precatheter
)

dev.off()


# ----------------------------------------------------------
# 8.19 控制台输出最终结果
# ----------------------------------------------------------

cat(
  "\nMICE多重插补全部完成。\n"
)

cat(
  "术前插补数据集数量：",
  imp_preoperative$m,
  "\n"
)

cat(
  "拔管前插补数据集数量：",
  imp_precatheter$m,
  "\n"
)

cat(
  "术前MICE logged events：",
  nrow(
    preoperative_logged_events
  ),
  "\n"
)

cat(
  "拔管前MICE logged events：",
  nrow(
    precatheter_logged_events
  ),
  "\n\n"
)

cat(
  "请检查：\n",
  "01_data/MICE_settings.xlsx\n",
  "05_logs/MICE_logged_events.xlsx\n",
  "05_logs/MICE_convergence_preoperative.pdf\n",
  "05_logs/MICE_convergence_precatheter.pdf\n"
)



############################################################
############################################################
############################################################
############################################################
############################################################
############################################################


cat("\n>>> 第9部分开始执行 <<<\n")


# 9.1 偏度
skewness_simple <- function(x) {
  x <- x[!is.na(x) & is.finite(x)]
  if (length(x) < 3 || length(unique(x)) < 2 || sd(x) == 0) return(0)
  mean((x - mean(x))^3) / sd(x)^3
}


# 9.2 Shapiro-Wilk P值，仅作审计
shapiro_simple <- function(x) {
  x <- x[!is.na(x) & is.finite(x)]
  if (length(x) < 3 || length(unique(x)) < 3) return(NA_real_)
  tryCatch(stats::shapiro.test(x)$p.value, error = function(e) NA_real_)
}


# 9.3 单个连续变量
summarise_continuous <- function(data, v) {
  
  x <- data[[v]]
  y <- data[[outcome_var]]
  
  xa <- x[!is.na(x)]
  xp <- x[y == 1 & !is.na(x)]
  xn <- x[y == 0 & !is.na(x)]
  
  skew_p <- skewness_simple(xp)
  skew_n <- skewness_simple(xn)
  normal_flag <- abs(skew_p) <= 1 && abs(skew_n) <= 1
  
  if (normal_flag) {
    test <- tryCatch(stats::t.test(xp, xn, var.equal = FALSE),
                     error = function(e) NULL)
    
    overall <- sprintf("%.2f ± %.2f", mean(xa), sd(xa))
    pour <- sprintf("%.2f ± %.2f", mean(xp), sd(xp))
    nonpour <- sprintf("%.2f ± %.2f", mean(xn), sd(xn))
    
    if (is.null(test)) {
      stat <- ""
      p <- NA_real_
      method <- "Welch t-test failed"
    } else {
      stat <- sprintf("t = %.3f", unname(test$statistic))
      p <- test$p.value
      method <- "Welch t-test"
    }
    
  } else {
    qa <- stats::quantile(xa, c(.25, .50, .75), na.rm = TRUE)
    qp <- stats::quantile(xp, c(.25, .50, .75), na.rm = TRUE)
    qn <- stats::quantile(xn, c(.25, .50, .75), na.rm = TRUE)
    
    test <- tryCatch(stats::wilcox.test(xp, xn, exact = FALSE),
                     error = function(e) NULL)
    
    overall <- sprintf("%.2f (%.2f, %.2f)", qa[2], qa[1], qa[3])
    pour <- sprintf("%.2f (%.2f, %.2f)", qp[2], qp[1], qp[3])
    nonpour <- sprintf("%.2f (%.2f, %.2f)", qn[2], qn[1], qn[3])
    
    if (is.null(test)) {
      stat <- ""
      p <- NA_real_
      method <- "Mann-Whitney U test failed"
    } else {
      stat <- sprintf("W = %.3f", unname(test$statistic))
      p <- test$p.value
      method <- "Mann-Whitney U test"
    }
  }
  
  display <- data.frame(
    Variable = v,
    Level = "",
    Overall = overall,
    POUR = pour,
    `Non-POUR` = nonpour,
    Statistic = stat,
    `P value` = format_p_value(p),
    check.names = FALSE
  )
  
  screening <- data.frame(
    Variable = v,
    P_value = p,
    Method = method
  )
  
  audit <- data.frame(
    Variable = v,
    Type = "Continuous",
    Observed_n = length(xa),
    Missing_n = sum(is.na(x)),
    Skewness_POUR = skew_p,
    Skewness_NonPOUR = skew_n,
    Shapiro_P_POUR = shapiro_simple(xp),
    Shapiro_P_NonPOUR = shapiro_simple(xn),
    Approx_normal = normal_flag,
    Method = method,
    Statistic = stat,
    Raw_P_value = p
  )
  
  list(display = display, screening = screening, audit = audit)
}


# 9.4 单个分类变量
summarise_categorical <- function(data, v, seed = 20260819) {
  
  x <- droplevels(factor(data[[v]]))
  y <- data[[outcome_var]]
  lev <- levels(x)
  
  p <- NA_real_
  stat <- ""
  method <- "Only one observed level"
  
  if (length(lev) >= 2) {
    
    tab <- table(x, y)
    
    expected <- tryCatch(
      suppressWarnings(stats::chisq.test(tab, correct = FALSE)$expected),
      error = function(e) NULL
    )
    
    sparse <- is.null(expected) || any(expected < 5)
    
    if (nrow(tab) == 2 && sparse) {
      test <- tryCatch(stats::fisher.test(tab), error = function(e) NULL)
      
      if (is.null(test)) {
        method <- "Fisher exact test failed"
      } else {
        p <- test$p.value
        stat <- "Fisher"
        method <- "Fisher exact test"
      }
      
    } else if (nrow(tab) > 2 && sparse) {
      set.seed(seed)
      
      test <- tryCatch(
        stats::chisq.test(tab, simulate.p.value = TRUE, B = 10000),
        error = function(e) NULL
      )
      
      if (is.null(test)) {
        method <- "Monte Carlo chi-square test failed"
      } else {
        p <- test$p.value
        stat <- sprintf("Chi-square = %.3f", unname(test$statistic))
        method <- "Pearson chi-square with Monte Carlo P (B=10000)"
      }
      
    } else {
      test <- tryCatch(
        stats::chisq.test(tab, correct = FALSE),
        error = function(e) NULL
      )
      
      if (is.null(test)) {
        method <- "Pearson chi-square test failed"
      } else {
        p <- test$p.value
        stat <- sprintf("Chi-square = %.3f", unname(test$statistic))
        method <- "Pearson chi-square test"
      }
    }
  }
  
  d_all <- sum(!is.na(data[[v]]))
  d_p <- sum(y == 1 & !is.na(data[[v]]))
  d_n <- sum(y == 0 & !is.na(data[[v]]))
  
  rows <- list()
  
  for (j in seq_along(lev)) {
    z <- lev[j]
    
    n_all <- sum(x == z, na.rm = TRUE)
    n_p <- sum(x == z & y == 1, na.rm = TRUE)
    n_n <- sum(x == z & y == 0, na.rm = TRUE)
    
    rows[[j]] <- data.frame(
      Variable = v,
      Level = z,
      Overall = sprintf("%d (%.1f%%)", n_all, 100 * n_all / d_all),
      POUR = sprintf("%d (%.1f%%)", n_p, 100 * n_p / d_p),
      `Non-POUR` = sprintf("%d (%.1f%%)", n_n, 100 * n_n / d_n),
      Statistic = ifelse(j == 1, stat, ""),
      `P value` = ifelse(j == 1, format_p_value(p), ""),
      check.names = FALSE
    )
  }
  
  display <- dplyr::bind_rows(rows)
  
  screening <- data.frame(
    Variable = v,
    P_value = p,
    Method = method
  )
  
  audit <- data.frame(
    Variable = v,
    Type = "Categorical",
    Observed_n = d_all,
    Missing_n = sum(is.na(data[[v]])),
    Skewness_POUR = NA_real_,
    Skewness_NonPOUR = NA_real_,
    Shapiro_P_POUR = NA_real_,
    Shapiro_P_NonPOUR = NA_real_,
    Approx_normal = NA,
    Method = method,
    Statistic = stat,
    Raw_P_value = p
  )
  
  list(display = display, screening = screening, audit = audit)
}


# 9.5 一个时间点的Table 1
make_table1_simple <- function(data, vars, name) {
  
  cat("\n开始：", name, "\n")
  
  displays <- list()
  screenings <- list()
  audits <- list()
  
  for (i in seq_along(vars)) {
    
    v <- vars[i]
    cat("  ", i, "/", length(vars), "  ", v, "\n", sep = "")
    flush.console()
    
    if (v %in% continuous_vars) {
      r <- summarise_continuous(data, v)
    } else {
      r <- summarise_categorical(data, v, seed = 20260819 + i)
    }
    
    displays[[i]] <- r$display
    screenings[[i]] <- r$screening
    audits[[i]] <- r$audit
  }
  
  cat("完成：", name, "\n")
  
  list(
    display = dplyr::bind_rows(displays),
    screening = dplyr::bind_rows(screenings),
    audit = dplyr::bind_rows(audits)
  )
}


cat(">>> 第9部分函数定义完成 <<<\n")


# 9.6 术前Table 1
preoperative_table1 <- make_table1_simple(
  analysis_data,
  preoperative_table_vars,
  "Preoperative Table 1"
)


# 9.7 拔管前Table 1
precatheter_table1 <- make_table1_simple(
  analysis_data,
  precatheter_table_vars,
  "Pre-catheter-removal Table 1"
)


# 9.8 保存正式Table 1
file_table1 <- file.path(
  table_dir,
  "Complete_cohort_characteristics.xlsx"
)

openxlsx::write.xlsx(
  list(
    Preoperative = preoperative_table1$display,
    Precatheter = precatheter_table1$display
  ),
  file_table1,
  overwrite = TRUE
)


# 9.9 保存P值筛选表
file_screen <- file.path(
  table_dir,
  "Univariable_screening_P_values.xlsx"
)

openxlsx::write.xlsx(
  list(
    Preoperative = preoperative_table1$screening,
    Precatheter = precatheter_table1$screening
  ),
  file_screen,
  overwrite = TRUE
)


# 9.10 保存统计方法审计表
file_audit <- file.path(
  log_dir,
  "Cohort_characteristics_statistical_method_audit.xlsx"
)

openxlsx::write.xlsx(
  list(
    Preoperative = preoperative_table1$audit,
    Precatheter = precatheter_table1$audit
  ),
  file_audit,
  overwrite = TRUE
)


# 9.11 检查文件是否真正生成
stopifnot(
  file.exists(file_table1),
  file.exists(file_screen),
  file.exists(file_audit)
)


# 9.12 显示P<0.10变量
cat("\n术前 P < 0.10：\n")
print(
  dplyr::filter(
    preoperative_table1$screening,
    !is.na(P_value) & P_value < 0.10
  )
)

cat("\n拔管前 P < 0.10：\n")
print(
  dplyr::filter(
    precatheter_table1$screening,
    !is.na(P_value) & P_value < 0.10
  )
)

cat("\n>>> 第9部分全部完成 <<<\n")
cat("Table 1：", file_table1, "\n")
cat("筛选P值：", file_screen, "\n")
cat("方法审计：", file_audit, "\n")




############################################################
# 10. 根据 Table 1 形成多因素候选变量池
############################################################


# ----------------------------------------------------------
# 10.1 从 Table 1 提取 P < 0.10 的变量
# ----------------------------------------------------------

preoperative_selected_table1 <- (
  preoperative_table1$screening %>%
    dplyr::filter(
      !is.na(P_value) &
        P_value < 0.10
    ) %>%
    dplyr::pull(Variable)
)


precatheter_selected_table1 <- (
  precatheter_table1$screening %>%
    dplyr::filter(
      !is.na(P_value) &
        P_value < 0.10
    ) %>%
    dplyr::pull(Variable)
)


# ----------------------------------------------------------
# 10.2 将原始POP分期变量替换为建模用分期变量
#
# Table 1：
# 0、I、II、III、IV
#
# Logistic建模：
# 0-I、II、III、IV
# ----------------------------------------------------------

replace_stage_for_model <- function(vars) {
  
  vars[
    vars == "uterine_prolapse_stage"
  ] <- "uterine_prolapse_stage_model"
  
  vars[
    vars == "anterior_wall_stage"
  ] <- "anterior_wall_stage_model"
  
  vars[
    vars == "posterior_wall_stage"
  ] <- "posterior_wall_stage_model"
  
  unique(vars)
}


preoperative_selected_model <- replace_stage_for_model(
  preoperative_selected_table1
)


precatheter_selected_model <- replace_stage_for_model(
  precatheter_selected_table1
)


# ----------------------------------------------------------
# 10.3 输出初始候选变量
# ----------------------------------------------------------

cat("\n术前多因素初始候选变量：\n")
print(
  preoperative_selected_model
)

cat("\n拔管前多因素初始候选变量：\n")
print(
  precatheter_selected_model
)


# ----------------------------------------------------------
# 10.4 检查可能存在同源信息的变量组
# ----------------------------------------------------------

check_related_groups <- function(
    selected_vars,
    related_groups
) {
  
  result <- data.frame(
    Group = character(),
    Selected_variables = character(),
    N_selected = integer(),
    Need_model_comparison = logical(),
    stringsAsFactors = FALSE
  )
  
  for (group_name in names(related_groups)) {
    
    group_vars <- related_groups[[group_name]]
    
    # POP原始分期名称转换为建模名称
    group_vars <- replace_stage_for_model(
      group_vars
    )
    
    selected_in_group <- intersect(
      selected_vars,
      group_vars
    )
    
    if (length(selected_in_group) > 0) {
      
      result <- rbind(
        result,
        data.frame(
          Group = group_name,
          Selected_variables = paste(
            selected_in_group,
            collapse = " | "
          ),
          N_selected = length(
            selected_in_group
          ),
          Need_model_comparison = (
            length(selected_in_group) >= 2
          ),
          stringsAsFactors = FALSE
        )
      )
    }
  }
  
  result
}


preoperative_related_check <- check_related_groups(
  selected_vars =
    preoperative_selected_model,
  related_groups =
    related_variable_groups
)


precatheter_related_check <- check_related_groups(
  selected_vars =
    precatheter_selected_model,
  related_groups =
    related_variable_groups
)


# ----------------------------------------------------------
# 10.5 保存候选变量与相关变量检查结果
# ----------------------------------------------------------

openxlsx::write.xlsx(
  list(
    Preoperative_selected =
      data.frame(
        Variable =
          preoperative_selected_model
      ),
    
    Precatheter_selected =
      data.frame(
        Variable =
          precatheter_selected_model
      ),
    
    Preoperative_related =
      preoperative_related_check,
    
    Precatheter_related =
      precatheter_related_check
  ),
  
  file.path(
    table_dir,
    "Multivariable_candidate_variables.xlsx"
  ),
  
  overwrite = TRUE
)


# ----------------------------------------------------------
# 10.6 控制台显示需要候选模型比较的变量组
# ----------------------------------------------------------

cat(
  "\n术前需要进行替代模型比较的变量组：\n"
)

print(
  preoperative_related_check %>%
    dplyr::filter(
      Need_model_comparison
    )
)


cat(
  "\n拔管前需要进行替代模型比较的变量组：\n"
)

print(
  precatheter_related_check %>%
    dplyr::filter(
      Need_model_comparison
    )
)


cat(
  "\n第10部分候选变量整理完成。\n"
)
############################################################
############################################################


############################################################
############################################################
############################################################
# 11. 最终候选变量路线锁定与MI建模准备质控
#
# 已完成的方法学裁决：
# 1. Table 1 中 P < 0.10 的变量形成候选变量池；
# 2. posterior_wall_stage_model 与 popq_bp 已完成比较；
# 3. 两个时间点均固定采用连续变量 popq_bp；
# 4. posterior_wall_stage_model 不再进入后续正式多因素模型；
# 5. 本节不再重复拟合 Stage vs Bp 候选模型；
# 6. 本节只做最终路线锁定及20个MI数据集的建模准备质控；
# 7. 本节不覆盖 preoperative_selected_model / precatheter_selected_model，
#    因此不会改变第12部分的输入对象；
# 8. 第12部分仍将按自身规则再次 setdiff(stage) 并确认 popq_bp，
#    因此本节与第12部分完全兼容。
############################################################

cat("\n>>> 第11部分开始执行：固定 Bp 路线 <<<\n")
flush.console()


# ----------------------------------------------------------
# 11.1 运行前检查
# ----------------------------------------------------------

required_objects_11 <- c(
  "imp_preoperative",
  "imp_precatheter",
  "preoperative_selected_model",
  "precatheter_selected_model",
  "outcome_var",
  "table_dir",
  "log_dir"
)

missing_objects_11 <- required_objects_11[
  !vapply(
    required_objects_11,
    exists,
    logical(1),
    inherits = TRUE
  )
]

if (length(missing_objects_11) > 0) {
  stop(
    paste0(
      "第11部分缺少前置对象：",
      paste(missing_objects_11, collapse = ", ")
    )
  )
}

required_packages_11 <- c(
  "mice",
  "dplyr",
  "openxlsx"
)

missing_packages_11 <- required_packages_11[
  !vapply(
    required_packages_11,
    requireNamespace,
    logical(1),
    quietly = TRUE
  )
]

if (length(missing_packages_11) > 0) {
  stop(
    paste0(
      "第11部分缺少R包：",
      paste(missing_packages_11, collapse = ", ")
    )
  )
}

if (imp_preoperative$m != 20) {
  stop("术前MICE对象不是20个插补数据集。")
}

if (imp_precatheter$m != 20) {
  stop("拔管前MICE对象不是20个插补数据集。")
}

dir.create(
  table_dir,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  log_dir,
  recursive = TRUE,
  showWarnings = FALSE
)

cat("11.1 运行前检查：通过。\n")


# ----------------------------------------------------------
# 11.2 锁定最终候选变量路线
#
# 重要：
# 不修改上游对象 preoperative_selected_model / precatheter_selected_model。
# 这里只建立第11部分自己的 route vars，供质控与审计使用。
#
# 第12部分使用完全相同的核心逻辑：
# setdiff(selected_model, "posterior_wall_stage_model")
# ----------------------------------------------------------

# 上游候选变量不应存在重复项
if (anyDuplicated(preoperative_selected_model) > 0) {
  stop("preoperative_selected_model 中存在重复变量。")
}

if (anyDuplicated(precatheter_selected_model) > 0) {
  stop("precatheter_selected_model 中存在重复变量。")
}

preoperative_route_vars_11 <- setdiff(
  preoperative_selected_model,
  "posterior_wall_stage_model"
)

precatheter_route_vars_11 <- setdiff(
  precatheter_selected_model,
  "posterior_wall_stage_model"
)


# 两个时间点均必须保留 popq_bp
if (!"popq_bp" %in% preoperative_route_vars_11) {
  stop("术前最终路线中缺少 popq_bp。")
}

if (!"popq_bp" %in% precatheter_route_vars_11) {
  stop("拔管前最终路线中缺少 popq_bp。")
}


# stage不得继续进入正式路线
if ("posterior_wall_stage_model" %in% preoperative_route_vars_11) {
  stop("术前最终路线仍残留 posterior_wall_stage_model。")
}

if ("posterior_wall_stage_model" %in% precatheter_route_vars_11) {
  stop("拔管前最终路线仍残留 posterior_wall_stage_model。")
}


cat("\n术前最终路线候选变量：\n")
print(preoperative_route_vars_11)

cat("\n拔管前最终路线候选变量：\n")
print(precatheter_route_vars_11)


# ----------------------------------------------------------
# 11.2A 候选变量数量质控
#
# 根据当前正式Table 1结果：
# 术前：8项
# 拔管前：11项
#
# 与第12部分保持一致：数量不符时warning，不直接终止。
# ----------------------------------------------------------

if (length(preoperative_route_vars_11) != 8) {
  warning(
    paste0(
      "术前最终路线当前为 ",
      length(preoperative_route_vars_11),
      " 个候选变量，而当前预期为8个，请核查Table 1。"
    )
  )
}

if (length(precatheter_route_vars_11) != 11) {
  warning(
    paste0(
      "拔管前最终路线当前为 ",
      length(precatheter_route_vars_11),
      " 个候选变量，而当前预期为11个，请核查Table 1。"
    )
  )
}


# ----------------------------------------------------------
# 11.3 单个MICE对象的建模准备质控函数
#
# 检查：
# 1. 20个complete数据集均存在所有建模变量；
# 2. outcome和预测变量均无缺失；
# 3. outcome为明确0/1；
# 4. 每个预测变量至少有2个实际取值；
# 5. factor/character预测变量若>2个实际水平，则提示第12部分
#    的1自由度建模规则可能不兼容；
# 6. popq_bp必须是numeric/integer，确保按连续变量进入模型。
# ----------------------------------------------------------

check_mi_model_readiness_11 <- function(
    imp_object,
    predictor_vars,
    timepoint_name
) {
  
  m <- imp_object$m
  
  dataset_qc_list <- vector("list", m)
  variable_qc_list <- vector("list", m)
  
  for (i in seq_len(m)) {
    
    cat(
      "  ",
      timepoint_name,
      "：检查插补数据集 ",
      i,
      "/",
      m,
      "\n",
      sep = ""
    )
    flush.console()
    
    data_i <- mice::complete(
      imp_object,
      action = i
    )
    
    required_vars_now <- c(
      outcome_var,
      predictor_vars
    )
    
    missing_vars_now <- setdiff(
      required_vars_now,
      names(data_i)
    )
    
    if (length(missing_vars_now) > 0) {
      stop(
        paste0(
          timepoint_name,
          " 第",
          i,
          "个插补数据集缺少变量：",
          paste(missing_vars_now, collapse = ", ")
        )
      )
    }
    
    if (anyNA(data_i[required_vars_now])) {
      stop(
        paste0(
          timepoint_name,
          " 第",
          i,
          "个插补数据集的建模变量仍存在缺失。"
        )
      )
    }
    
    y_character <- trimws(
      as.character(data_i[[outcome_var]])
    )
    
    y_values <- sort(
      unique(y_character)
    )
    
    if (!all(y_values %in% c("0", "1"))) {
      stop(
        paste0(
          timepoint_name,
          " 第",
          i,
          "个插补数据集的结局变量不是明确0/1编码；实际取值：",
          paste(y_values, collapse = ", ")
        )
      )
    }
    
    event_n <- sum(
      y_character == "1"
    )
    
    nonevent_n <- sum(
      y_character == "0"
    )
    
    if (event_n == 0 || nonevent_n == 0) {
      stop(
        paste0(
          timepoint_name,
          " 第",
          i,
          "个插补数据集缺少事件组或非事件组。"
        )
      )
    }
    
    # popq_bp必须按连续变量进入正式路线
    if (!is.numeric(data_i$popq_bp) && !is.integer(data_i$popq_bp)) {
      stop(
        paste0(
          timepoint_name,
          " 第",
          i,
          "个插补数据集中的 popq_bp 不是numeric/integer，",
          "无法保证第12部分按连续变量建模。"
        )
      )
    }
    
    dataset_qc_list[[i]] <- data.frame(
      Timepoint = timepoint_name,
      Imputation = i,
      N = nrow(data_i),
      Events = event_n,
      Non_events = nonevent_n,
      Predictor_n = length(predictor_vars),
      Missing_model_variables = 0,
      stringsAsFactors = FALSE
    )
    
    variable_rows_i <- lapply(
      predictor_vars,
      function(v) {
        
        x <- data_i[[v]]
        
        x_nonmissing <- x[!is.na(x)]
        
        unique_n <- length(
          unique(x_nonmissing)
        )
        
        variable_class <- paste(
          class(x),
          collapse = "/"
        )
        
        is_categorical <- is.factor(x) || is.character(x)
        
        actual_level_n <- if (is_categorical) {
          length(
            unique(
              as.character(x_nonmissing)
            )
          )
        } else {
          NA_integer_
        }
        
        one_df_compatible <- if (is_categorical) {
          actual_level_n <= 2
        } else {
          TRUE
        }
        
        data.frame(
          Timepoint = timepoint_name,
          Imputation = i,
          Variable = v,
          Class = variable_class,
          Unique_nonmissing_values = unique_n,
          Actual_categorical_levels = actual_level_n,
          One_df_compatible = one_df_compatible,
          Constant_variable = unique_n < 2,
          stringsAsFactors = FALSE
        )
      }
    )
    
    variable_qc_list[[i]] <- dplyr::bind_rows(
      variable_rows_i
    )
  }
  
  dataset_qc <- dplyr::bind_rows(
    dataset_qc_list
  )
  
  variable_qc <- dplyr::bind_rows(
    variable_qc_list
  )
  
  if (any(variable_qc$Constant_variable)) {
    bad_rows <- variable_qc[
      variable_qc$Constant_variable,
      ,
      drop = FALSE
    ]
    
    stop(
      paste0(
        timepoint_name,
        " 至少一个候选变量在某个插补数据集中为常数。请检查：",
        paste(
          unique(bad_rows$Variable),
          collapse = ", "
        )
      )
    )
  }
  
  if (any(!variable_qc$One_df_compatible)) {
    bad_rows <- variable_qc[
      !variable_qc$One_df_compatible,
      ,
      drop = FALSE
    ]
    
    stop(
      paste0(
        timepoint_name,
        " 存在>2实际水平的分类候选变量，不符合当前第12部分",
        "‘所有候选变量均为1自由度’的规则。请检查：",
        paste(
          unique(bad_rows$Variable),
          collapse = ", "
        )
      )
    )
  }
  
  list(
    dataset_qc = dataset_qc,
    variable_qc = variable_qc
  )
}


# ----------------------------------------------------------
# 11.4 对两个时间点完成20个MI数据集质控
# ----------------------------------------------------------

cat("\n开始术前MI建模准备质控：\n")

preoperative_readiness_11 <- check_mi_model_readiness_11(
  imp_object = imp_preoperative,
  predictor_vars = preoperative_route_vars_11,
  timepoint_name = "Preoperative"
)

cat("\n开始拔管前MI建模准备质控：\n")

precatheter_readiness_11 <- check_mi_model_readiness_11(
  imp_object = imp_precatheter,
  predictor_vars = precatheter_route_vars_11,
  timepoint_name = "Precatheter"
)

cat("11.4 两个时间点20个MI数据集建模准备质控：通过。\n")


# ----------------------------------------------------------
# 11.5 保存最终路线变量清单
# ----------------------------------------------------------

route_variable_table_11 <- dplyr::bind_rows(
  data.frame(
    Timepoint = "Preoperative",
    Variable = preoperative_route_vars_11,
    Route_decision = "Retained for Section 12",
    stringsAsFactors = FALSE
  ),
  data.frame(
    Timepoint = "Precatheter",
    Variable = precatheter_route_vars_11,
    Route_decision = "Retained for Section 12",
    stringsAsFactors = FALSE
  )
)

route_decision_table_11 <- data.frame(
  Decision = c(
    "Posterior wall representation",
    "Variable retained",
    "Variable excluded from formal route",
    "Reason for no repeated A/B modelling"
  ),
  Final_value = c(
    "Continuous POP-Q Bp",
    "popq_bp",
    "posterior_wall_stage_model",
    "Stage vs Bp comparison already completed and Bp route fixed"
  ),
  stringsAsFactors = FALSE
)

openxlsx::write.xlsx(
  list(
    Route_decision = route_decision_table_11,
    Final_candidate_variables = route_variable_table_11
  ),
  file.path(
    table_dir,
    "Section11_final_Bp_route_lock.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 11.6 保存MI建模准备质控
# ----------------------------------------------------------

mi_dataset_qc_11 <- dplyr::bind_rows(
  preoperative_readiness_11$dataset_qc,
  precatheter_readiness_11$dataset_qc
)

mi_variable_qc_11 <- dplyr::bind_rows(
  preoperative_readiness_11$variable_qc,
  precatheter_readiness_11$variable_qc
)

openxlsx::write.xlsx(
  list(
    Dataset_QC = mi_dataset_qc_11,
    Variable_QC = mi_variable_qc_11
  ),
  file.path(
    log_dir,
    "Section11_Bp_route_MI_readiness_QC.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 11.7 与第12部分输入逻辑进行一致性核验
#
# 这里明确模拟第12部分12.2的变量构建方式，
# 确保本节锁定的变量集合与第12部分实际使用的集合完全一致。
# ----------------------------------------------------------

preoperative_vars_as_section12 <- setdiff(
  preoperative_selected_model,
  "posterior_wall_stage_model"
)

precatheter_vars_as_section12 <- setdiff(
  precatheter_selected_model,
  "posterior_wall_stage_model"
)

if (!identical(
  preoperative_route_vars_11,
  preoperative_vars_as_section12
)) {
  stop("术前第11部分路线与第12部分变量构建逻辑不一致。")
}

if (!identical(
  precatheter_route_vars_11,
  precatheter_vars_as_section12
)) {
  stop("拔管前第11部分路线与第12部分变量构建逻辑不一致。")
}

cat("11.7 与第12部分输入逻辑一致性核验：通过。\n")


# ----------------------------------------------------------
# 11.8 最终文件质控
# ----------------------------------------------------------

required_output_files_11 <- c(
  file.path(
    table_dir,
    "Section11_final_Bp_route_lock.xlsx"
  ),
  file.path(
    log_dir,
    "Section11_Bp_route_MI_readiness_QC.xlsx"
  )
)

missing_output_files_11 <- required_output_files_11[
  !file.exists(required_output_files_11)
]

if (length(missing_output_files_11) > 0) {
  stop(
    paste0(
      "第11部分以下输出文件未成功生成：",
      paste(
        basename(missing_output_files_11),
        collapse = ", "
      )
    )
  )
}


# ----------------------------------------------------------
# 11.9 控制台输出
# ----------------------------------------------------------

cat("\n最终后壁脱垂表示方式：popq_bp（连续变量）\n")

cat("\n术前第12部分将使用的完整候选变量：\n")
print(preoperative_route_vars_11)

cat("\n拔管前第12部分将使用的完整候选变量：\n")
print(precatheter_route_vars_11)

cat(
  "\n>>> 第11部分全部完成：Bp路线已锁定，可直接运行第12部分 <<<\n"
)

cat(
  "请检查：\n",
  "02_tables/Section11_final_Bp_route_lock.xlsx\n",
  "05_logs/Section11_Bp_route_MI_readiness_QC.xlsx\n"
)


############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
# 12. 最终多因素 Logistic 模型筛选、完整/最终模型输出及表观性能
#
# 已完成的前置裁决：
# 1. Table 1 中 P < 0.10 形成候选变量池；
# 2. posterior_wall_stage_model 与 popq_bp 比较后，
#    两个时间点均固定采用 popq_bp；
# 3. 本节不再使用 posterior_wall_stage_model；
# 4. 在20个MICE数据集中分别拟合模型，
#    β、SE、P值和95%CI使用Rubin's rules合并；
# 5. 当前候选变量全部为1自由度变量，
#    因此完整模型中 pooled P < 0.05 的变量进入最终模型；
# 6. 最终模型重新拟合后生成完整/最终多因素模型输出；
# 7. 同时输出表观AUC、Brier、Scaled Brier/IPA、Log loss、AIC/BIC。
############################################################

cat("\n>>> 第12部分开始执行 <<<\n")
flush.console()


# ----------------------------------------------------------
# 12.1 运行前检查
# ----------------------------------------------------------

required_objects_12 <- c(
  "imp_preoperative",
  "imp_precatheter",
  "preoperative_selected_model",
  "precatheter_selected_model",
  "outcome_var",
  "table_dir",
  "model_dir",
  "log_dir",
  "format_p_value"
)

missing_objects_12 <- required_objects_12[
  !vapply(
    required_objects_12,
    exists,
    logical(1),
    inherits = TRUE
  )
]

if (length(missing_objects_12) > 0) {
  
  stop(
    paste0(
      "第12部分缺少前置对象：",
      paste(
        missing_objects_12,
        collapse = ", "
      )
    )
  )
}

if (imp_preoperative$m != 20) {
  stop("术前MICE对象不是20个插补数据集。")
}

if (imp_precatheter$m != 20) {
  stop("拔管前MICE对象不是20个插补数据集。")
}

cat("12.1 运行前检查：通过。\n")


# ----------------------------------------------------------
# 12.2 锁定两个时间点的完整多因素候选变量
#
# posterior_wall_stage_model 已正式舍弃，
# 两个时间点均采用 popq_bp。
# ----------------------------------------------------------

preoperative_full_vars <- setdiff(
  preoperative_selected_model,
  "posterior_wall_stage_model"
)

precatheter_full_vars <- setdiff(
  precatheter_selected_model,
  "posterior_wall_stage_model"
)


# 必须保留 popq_bp
if (!"popq_bp" %in% preoperative_full_vars) {
  stop("术前完整模型中缺少 popq_bp。")
}

if (!"popq_bp" %in% precatheter_full_vars) {
  stop("拔管前完整模型中缺少 popq_bp。")
}


# 不允许 stage 继续残留
if (
  any(
    grepl(
      "posterior_wall_stage_model",
      preoperative_full_vars
    )
  )
) {
  stop("术前变量池仍残留 posterior_wall_stage_model。")
}

if (
  any(
    grepl(
      "posterior_wall_stage_model",
      precatheter_full_vars
    )
  )
) {
  stop("拔管前变量池仍残留 posterior_wall_stage_model。")
}


cat("\n术前完整模型候选变量：\n")
print(preoperative_full_vars)

cat("\n拔管前完整模型候选变量：\n")
print(precatheter_full_vars)


# ----------------------------------------------------------
# 12.2A 候选变量数量质控
#
# 根据当前Table 1结果：
# 术前应为8项
# 拔管前应为11项
# ----------------------------------------------------------

if (length(preoperative_full_vars) != 8) {
  
  warning(
    paste0(
      "术前完整模型当前为 ",
      length(preoperative_full_vars),
      " 个候选变量，而预期为8个，请核查Table 1筛选结果。"
    )
  )
}

if (length(precatheter_full_vars) != 11) {
  
  warning(
    paste0(
      "拔管前完整模型当前为 ",
      length(precatheter_full_vars),
      " 个候选变量，而预期为11个，请核查Table 1筛选结果。"
    )
  )
}


full_candidate_table <- dplyr::bind_rows(
  
  data.frame(
    Timepoint = "Preoperative",
    Variable = preoperative_full_vars,
    stringsAsFactors = FALSE
  ),
  
  data.frame(
    Timepoint = "Precatheter",
    Variable = precatheter_full_vars,
    stringsAsFactors = FALSE
  )
)

openxlsx::write.xlsx(
  full_candidate_table,
  file.path(
    table_dir,
    "Final_route_full_model_candidate_variables.xlsx"
  ),
  overwrite = TRUE
)

cat("12.2 完整多因素候选变量锁定：完成。\n")


# ----------------------------------------------------------
# 12.3 根据一个glm模型建立“变量-系数项”映射
#
# 当前路线要求所有候选预测变量均为1自由度：
# - 连续变量 → 1个系数
# - 二分类factor → 1个系数
#
# 如果某变量产生>1个系数，本节直接停止，
# 防止错误使用单个dummy P值代替整体变量P值。
# ----------------------------------------------------------

make_one_df_term_map <- function(
    fit,
    predictor_vars
) {
  
  mm <- stats::model.matrix(fit)
  
  assign_vector <- attr(
    mm,
    "assign"
  )
  
  term_labels <- attr(
    stats::terms(fit),
    "term.labels"
  )
  
  mm_columns <- colnames(mm)
  
  result_list <- list()
  
  
  for (i in seq_along(term_labels)) {
    
    variable_name <- term_labels[i]
    
    coefficient_terms <- mm_columns[
      assign_vector == i
    ]
    
    if (length(coefficient_terms) != 1) {
      
      stop(
        paste0(
          "变量 ",
          variable_name,
          " 产生了 ",
          length(coefficient_terms),
          " 个回归系数；",
          "当前第12部分仅允许1自由度变量。",
          "如出现多分类变量，应改用MI整体变量检验。"
        )
      )
    }
    
    result_list[[i]] <- data.frame(
      Variable = variable_name,
      Term = coefficient_terms,
      stringsAsFactors = FALSE
    )
  }
  
  
  result <- dplyr::bind_rows(
    result_list
  )
  
  
  missing_predictors <- setdiff(
    predictor_vars,
    result$Variable
  )
  
  if (length(missing_predictors) > 0) {
    
    stop(
      paste0(
        "变量-系数映射缺少：",
        paste(
          missing_predictors,
          collapse = ", "
        )
      )
    )
  }
  
  
  result
}


# ----------------------------------------------------------
# 12.4 VIF/GVIF函数
# ----------------------------------------------------------

extract_vif_12 <- function(
    fit,
    model_name,
    imputation_id
) {
  
  model_terms <- attr(
    stats::terms(fit),
    "term.labels"
  )
  
  
  # 截距模型
  if (length(model_terms) == 0) {
    
    return(
      data.frame(
        Model = model_name,
        Imputation = imputation_id,
        Variable = character(),
        Df = numeric(),
        VIF = numeric(),
        GVIF = numeric(),
        Adjusted_GVIF = numeric(),
        VIF_equivalent = numeric(),
        stringsAsFactors = FALSE
      )
    )
  }
  
  
  # 只有一个预测因子时，VIF定义为1
  if (length(model_terms) == 1) {
    
    return(
      data.frame(
        Model = model_name,
        Imputation = imputation_id,
        Variable = model_terms,
        Df = 1,
        VIF = 1,
        GVIF = NA_real_,
        Adjusted_GVIF = NA_real_,
        VIF_equivalent = 1,
        stringsAsFactors = FALSE
      )
    )
  }
  
  
  vif_object <- tryCatch(
    car::vif(fit),
    error = function(e) e
  )
  
  
  if (inherits(vif_object, "error")) {
    
    return(
      data.frame(
        Model = model_name,
        Imputation = imputation_id,
        Variable = "VIF calculation failed",
        Df = NA_real_,
        VIF = NA_real_,
        GVIF = NA_real_,
        Adjusted_GVIF = NA_real_,
        VIF_equivalent = NA_real_,
        stringsAsFactors = FALSE
      )
    )
  }
  
  
  # 多自由度factor
  if (is.matrix(vif_object)) {
    
    adjusted_name <- grep(
      "GVIF\\^",
      colnames(vif_object),
      value = TRUE
    )
    
    if (length(adjusted_name) >= 1) {
      
      adjusted_value <- vif_object[
        ,
        adjusted_name[1]
      ]
      
    } else {
      
      adjusted_value <- vif_object[
        ,
        "GVIF"
      ]^(
        1 /
          (
            2 *
              vif_object[
                ,
                "Df"
              ]
          )
      )
    }
    
    
    return(
      data.frame(
        Model = model_name,
        Imputation = imputation_id,
        Variable = rownames(vif_object),
        Df = vif_object[, "Df"],
        VIF = NA_real_,
        GVIF = vif_object[, "GVIF"],
        Adjusted_GVIF = adjusted_value,
        VIF_equivalent = adjusted_value^2,
        stringsAsFactors = FALSE
      )
    )
  }
  
  
  # 普通VIF
  data.frame(
    Model = model_name,
    Imputation = imputation_id,
    Variable = names(vif_object),
    Df = 1,
    VIF = as.numeric(vif_object),
    GVIF = NA_real_,
    Adjusted_GVIF = NA_real_,
    VIF_equivalent = as.numeric(vif_object),
    stringsAsFactors = FALSE
  )
}


# ----------------------------------------------------------
# 12.5 在一个MICE对象上拟合一个多因素Logistic模型
#
# 输出：
# - 20个glm对象
# - Rubin pooled结果
# - β、SE、Wald、P、OR及95%CI
# - VIF
# - 收敛/稀疏数据质控
# - 表观AIC/BIC/AUC，仅作内部质控
# ----------------------------------------------------------

fit_mi_logistic_12 <- function(
    imp_object,
    predictor_vars,
    model_name
) {
  
  m <- imp_object$m
  
  if (length(predictor_vars) == 0) {
    
    formula_object <- stats::as.formula(
      paste0(
        outcome_var,
        " ~ 1"
      )
    )
    
  } else {
    
    formula_object <- stats::as.formula(
      paste(
        outcome_var,
        "~",
        paste(
          predictor_vars,
          collapse = " + "
        )
      )
    )
  }
  
  
  fit_list <- vector(
    "list",
    m
  )
  
  qc_list <- vector(
    "list",
    m
  )
  
  vif_list <- vector(
    "list",
    m
  )
  
  
  cat(
    "\n开始拟合：",
    model_name,
    "\n"
  )
  
  
  for (i in seq_len(m)) {
    
    cat(
      "  插补数据集 ",
      i,
      "/",
      m,
      "\n",
      sep = ""
    )
    flush.console()
    
    
    data_i <- mice::complete(
      imp_object,
      action = i
    )
    
    
    required_vars_now <- c(
      outcome_var,
      predictor_vars
    )
    
    
    missing_vars_now <- setdiff(
      required_vars_now,
      names(data_i)
    )
    
    if (length(missing_vars_now) > 0) {
      
      stop(
        paste0(
          model_name,
          " 第",
          i,
          "个插补数据集缺少变量：",
          paste(
            missing_vars_now,
            collapse = ", "
          )
        )
      )
    }
    
    
    if (
      anyNA(
        data_i[
          required_vars_now
        ]
      )
    ) {
      
      stop(
        paste0(
          model_name,
          " 第",
          i,
          "个插补数据集仍存在建模变量缺失。"
        )
      )
    }
    
    
    captured_warnings <- character()
    
    
    fit_i <- tryCatch(
      
      withCallingHandlers(
        
        stats::glm(
          formula = formula_object,
          data = data_i,
          family = stats::binomial()
        ),
        
        warning = function(w) {
          
          captured_warnings <<- c(
            captured_warnings,
            conditionMessage(w)
          )
          
          invokeRestart(
            "muffleWarning"
          )
        }
      ),
      
      error = function(e) e
    )
    
    
    if (inherits(fit_i, "error")) {
      
      stop(
        paste0(
          model_name,
          " 第",
          i,
          "个插补数据集拟合失败：",
          conditionMessage(fit_i)
        )
      )
    }
    
    
    fit_list[[i]] <- fit_i
    
    
    coef_table <- summary(
      fit_i
    )$coefficients
    
    
    extreme_coefficient <- any(
      abs(
        coef_table[
          ,
          "Estimate"
        ]
      ) > 10,
      na.rm = TRUE
    )
    
    
    very_large_se <- any(
      coef_table[
        ,
        "Std. Error"
      ] > 5,
      na.rm = TRUE
    )
    
    
    parameter_n <- sum(
      !is.na(
        stats::coef(fit_i)
      )
    ) - 1
    
    
    # ------------------------------------------------------
    # 将结局安全转换为0/1，用于模型性能计算
    # ------------------------------------------------------
    
    y_i <- data_i[[outcome_var]]
    
    if (is.factor(y_i) || is.character(y_i)) {
      y_i <- as.numeric(as.character(y_i))
    } else {
      y_i <- as.numeric(y_i)
    }
    
    if (anyNA(y_i) || !all(y_i %in% c(0, 1))) {
      stop(
        paste0(
          model_name,
          " 第",
          i,
          "个插补数据集的结局变量不是明确的0/1编码。"
        )
      )
    }
    
    
    event_n <- sum(
      y_i == 1
    )
    
    
    epv <- ifelse(
      parameter_n > 0,
      event_n /
        parameter_n,
      NA_real_
    )
    
    
    pred_i <- stats::predict(
      fit_i,
      type = "response"
    )
    
    
    roc_i <- tryCatch(
      
      pROC::roc(
        response = y_i,
        predictor = pred_i,
        levels = c(0, 1),
        direction = "<",
        quiet = TRUE
      ),
      
      error = function(e) NULL
    )
    
    
    apparent_auc <- if (
      is.null(roc_i)
    ) {
      NA_real_
    } else {
      as.numeric(
        pROC::auc(
          roc_i
        )
      )
    }
    
    
    # ------------------------------------------------------
    # 表观模型性能
    #
    # 注意：以下性能均在“拟合该模型的同一插补数据集”上计算，
    # 因此属于 apparent performance，而不是内部验证性能。
    # ------------------------------------------------------
    
    brier_score_i <- mean(
      (y_i - pred_i)^2
    )
    
    event_rate_i <- mean(
      y_i
    )
    
    null_brier_i <- mean(
      (y_i - event_rate_i)^2
    )
    
    scaled_brier_ipa_i <- if (
      is.finite(null_brier_i) &&
      null_brier_i > 0
    ) {
      1 - brier_score_i / null_brier_i
    } else {
      NA_real_
    }
    
    epsilon_12 <- 1e-15
    
    pred_i_clipped <- pmin(
      pmax(pred_i, epsilon_12),
      1 - epsilon_12
    )
    
    log_loss_i <- -mean(
      y_i * log(pred_i_clipped) +
        (1 - y_i) * log(1 - pred_i_clipped)
    )
    
    mean_predicted_risk_i <- mean(
      pred_i
    )
    
    
    qc_list[[i]] <- data.frame(
      Model = model_name,
      Imputation = i,
      N = nrow(data_i),
      Events = event_n,
      Parameters_without_intercept = parameter_n,
      EPV = epv,
      Converged = isTRUE(
        fit_i$converged
      ),
      Extreme_coefficient =
        extreme_coefficient,
      Very_large_SE =
        very_large_se,
      Warning_n =
        length(
          captured_warnings
        ),
      Warning = if (
        length(captured_warnings) == 0
      ) {
        ""
      } else {
        paste(
          unique(
            captured_warnings
          ),
          collapse = " | "
        )
      },
      AIC = stats::AIC(
        fit_i
      ),
      BIC = stats::BIC(
        fit_i
      ),
      Apparent_AUC =
        apparent_auc,
      Brier_Score =
        brier_score_i,
      Scaled_Brier_IPA =
        scaled_brier_ipa_i,
      Log_Loss =
        log_loss_i,
      Observed_Event_Rate =
        event_rate_i,
      Mean_Predicted_Risk =
        mean_predicted_risk_i,
      stringsAsFactors = FALSE
    )
    
    
    vif_list[[i]] <- extract_vif_12(
      fit = fit_i,
      model_name = model_name,
      imputation_id = i
    )
  }
  
  
  # --------------------------------------------------------
  # 变量-系数映射
  # --------------------------------------------------------
  
  if (length(predictor_vars) > 0) {
    
    term_map <- make_one_df_term_map(
      fit = fit_list[[1]],
      predictor_vars = predictor_vars
    )
    
  } else {
    
    term_map <- data.frame(
      Variable = character(),
      Term = character(),
      stringsAsFactors = FALSE
    )
  }
  
  
  # --------------------------------------------------------
  # Rubin pooling
  # --------------------------------------------------------
  
  mira_object <- mice::as.mira(
    fit_list
  )
  
  pooled_object <- mice::pool(
    mira_object
  )
  
  pooled_summary <- summary(
    pooled_object,
    conf.int = TRUE,
    conf.level = 0.95
  )
  
  
  # --------------------------------------------------------
  # 整理正式回归结果
  # --------------------------------------------------------
  
  pooled_table <- data.frame(
    Term = pooled_summary$term,
    Beta = pooled_summary$estimate,
    SE = pooled_summary$std.error,
    Statistic = pooled_summary$statistic,
    Df = pooled_summary$df,
    `P value` = pooled_summary$p.value,
    `CI lower beta` =
      pooled_summary$`2.5 %`,
    `CI upper beta` =
      pooled_summary$`97.5 %`,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
  
  
  pooled_table$Wald <- (
    pooled_table$Beta /
      pooled_table$SE
  )^2
  
  
  pooled_table$OR <- exp(
    pooled_table$Beta
  )
  
  pooled_table$`OR CI lower` <- exp(
    pooled_table$
      `CI lower beta`
  )
  
  pooled_table$`OR CI upper` <- exp(
    pooled_table$
      `CI upper beta`
  )
  
  
  pooled_table$`OR (95% CI)` <- sprintf(
    "%.3f (%.3f-%.3f)",
    pooled_table$OR,
    pooled_table$
      `OR CI lower`,
    pooled_table$
      `OR CI upper`
  )
  
  
  pooled_table$`P formatted` <- vapply(
    pooled_table$`P value`,
    format_p_value,
    character(1)
  )
  
  
  # --------------------------------------------------------
  # 加入原始变量名
  # --------------------------------------------------------
  
  pooled_table$Variable <- ifelse(
    pooled_table$Term ==
      "(Intercept)",
    "(Intercept)",
    NA_character_
  )
  
  
  if (nrow(term_map) > 0) {
    
    for (
      j in seq_len(
        nrow(term_map)
      )
    ) {
      
      pooled_table$Variable[
        pooled_table$Term ==
          term_map$Term[j]
      ] <- term_map$Variable[j]
    }
  }
  
  
  # --------------------------------------------------------
  # 变量映射质控
  # --------------------------------------------------------
  
  non_intercept_table <- pooled_table[
    pooled_table$Term !=
      "(Intercept)",
    ,
    drop = FALSE
  ]
  
  
  if (
    length(predictor_vars) > 0 &&
    any(
      is.na(
        non_intercept_table$
        Variable
      )
    )
  ) {
    
    stop(
      paste0(
        model_name,
        " pooled结果中存在无法映射回原始变量的系数。"
      )
    )
  }
  
  
  if (
    nrow(
      non_intercept_table
    ) !=
    length(
      predictor_vars
    )
  ) {
    
    stop(
      paste0(
        model_name,
        " pooled系数数目与候选变量数不一致。"
      )
    )
  }
  
  
  # --------------------------------------------------------
  # VIF与QC汇总
  # --------------------------------------------------------
  
  qc_all <- dplyr::bind_rows(
    qc_list
  )
  
  vif_all <- dplyr::bind_rows(
    vif_list
  )
  
  
  vif_summary <- if (
    nrow(vif_all) == 0
  ) {
    
    data.frame(
      Model = model_name,
      Variable = character(),
      Mean_VIF_equivalent = numeric(),
      Maximum_VIF_equivalent = numeric(),
      stringsAsFactors = FALSE
    )
    
  } else {
    
    vif_all %>%
      dplyr::group_by(
        Model,
        Variable
      ) %>%
      dplyr::summarise(
        Mean_VIF_equivalent =
          mean(
            VIF_equivalent,
            na.rm = TRUE
          ),
        Maximum_VIF_equivalent =
          max(
            VIF_equivalent,
            na.rm = TRUE
          ),
        .groups = "drop"
      )
  }
  
  
  qc_summary <- data.frame(
    Model = model_name,
    Number_of_fits = nrow(
      qc_all
    ),
    All_converged = all(
      qc_all$Converged
    ),
    Any_extreme_coefficient = any(
      qc_all$
        Extreme_coefficient
    ),
    Any_very_large_SE = any(
      qc_all$
        Very_large_SE
    ),
    Total_warnings = sum(
      qc_all$
        Warning_n
    ),
    Mean_EPV = if (
      all(
        is.na(
          qc_all$EPV
        )
      )
    ) {
      NA_real_
    } else {
      mean(
        qc_all$EPV,
        na.rm = TRUE
      )
    },
    Minimum_EPV = if (
      all(
        is.na(
          qc_all$EPV
        )
      )
    ) {
      NA_real_
    } else {
      min(
        qc_all$EPV,
        na.rm = TRUE
      )
    },
    Mean_AIC = mean(
      qc_all$AIC,
      na.rm = TRUE
    ),
    Mean_BIC = mean(
      qc_all$BIC,
      na.rm = TRUE
    ),
    Mean_Apparent_AUC = mean(
      qc_all$Apparent_AUC,
      na.rm = TRUE
    ),
    SD_Apparent_AUC = stats::sd(
      qc_all$Apparent_AUC,
      na.rm = TRUE
    ),
    Mean_Brier = mean(
      qc_all$Brier_Score,
      na.rm = TRUE
    ),
    SD_Brier = stats::sd(
      qc_all$Brier_Score,
      na.rm = TRUE
    ),
    Mean_Scaled_Brier_IPA = mean(
      qc_all$Scaled_Brier_IPA,
      na.rm = TRUE
    ),
    Mean_Log_Loss = mean(
      qc_all$Log_Loss,
      na.rm = TRUE
    ),
    Mean_Observed_Event_Rate = mean(
      qc_all$Observed_Event_Rate,
      na.rm = TRUE
    ),
    Mean_Predicted_Risk = mean(
      qc_all$Mean_Predicted_Risk,
      na.rm = TRUE
    ),
    Maximum_VIF_equivalent = if (
      nrow(vif_all) == 0 ||
      all(
        is.na(
          vif_all$
          VIF_equivalent
        )
      )
    ) {
      NA_real_
    } else {
      max(
        vif_all$
          VIF_equivalent,
        na.rm = TRUE
      )
    },
    stringsAsFactors = FALSE
  )
  
  
  cat(
    "完成：",
    model_name,
    "\n"
  )
  
  
  list(
    model_name = model_name,
    predictor_vars = predictor_vars,
    formula = formula_object,
    fits = fit_list,
    mira = mira_object,
    pooled = pooled_object,
    pooled_table = pooled_table,
    term_map = term_map,
    qc_detail = qc_all,
    qc_summary = qc_summary,
    vif_detail = vif_all,
    vif_summary = vif_summary
  )
}


cat("12.5 多重插补Logistic拟合函数定义完成。\n")


# ----------------------------------------------------------
# 12.6 拟合两个完整多因素模型
# ----------------------------------------------------------

preoperative_full_model_12 <- fit_mi_logistic_12(
  imp_object = imp_preoperative,
  predictor_vars = preoperative_full_vars,
  model_name = "Preoperative_full"
)


precatheter_full_model_12 <- fit_mi_logistic_12(
  imp_object = imp_precatheter,
  predictor_vars = precatheter_full_vars,
  model_name = "Precatheter_full"
)


# ----------------------------------------------------------
# 12.6A 完整模型质控
# ----------------------------------------------------------

full_model_qc <- dplyr::bind_rows(
  preoperative_full_model_12$
    qc_summary,
  precatheter_full_model_12$
    qc_summary
)


cat(
  "\n完整模型质控：\n"
)

print(
  full_model_qc
)


if (
  any(
    full_model_qc$
    Number_of_fits != 20
  )
) {
  stop(
    "完整模型没有全部完成20次拟合。"
  )
}


if (
  any(
    !full_model_qc$
    All_converged
  )
) {
  warning(
    "至少一个完整模型存在未收敛拟合。"
  )
}


if (
  any(
    full_model_qc$
    Any_extreme_coefficient
  )
) {
  warning(
    "至少一个完整模型出现|Beta|>10，请检查稀疏数据或分离。"
  )
}


if (
  any(
    full_model_qc$
    Any_very_large_SE
  )
) {
  warning(
    "至少一个完整模型出现SE>5，请检查稀疏数据或分离。"
  )
}


if (
  any(
    full_model_qc$
    Total_warnings > 0
  )
) {
  warning(
    "完整模型拟合产生warning，请检查质控文件。"
  )
}


if (
  any(
    full_model_qc$
    Maximum_VIF_equivalent >= 5,
    na.rm = TRUE
  )
) {
  warning(
    "完整模型存在VIF equivalent >=5，请检查共线性。"
  )
}


# ----------------------------------------------------------
# 12.7 提取完整模型每个变量的 pooled P 值
#
# 当前候选预测变量全部为1自由度，
# 因此每个变量对应1个pooled回归系数。
# ----------------------------------------------------------

extract_variable_selection_12 <- function(
    model_object,
    cutoff = 0.05
) {
  
  table_now <- model_object$
    pooled_table
  
  table_now <- table_now[
    table_now$Variable !=
      "(Intercept)",
    ,
    drop = FALSE
  ]
  
  
  result <- data.frame(
    Variable =
      table_now$Variable,
    Beta =
      table_now$Beta,
    SE =
      table_now$SE,
    Wald =
      table_now$Wald,
    P_value =
      table_now$
      `P value`,
    P_formatted =
      table_now$
      `P formatted`,
    OR =
      table_now$OR,
    OR_95CI =
      table_now$
      `OR (95% CI)`,
    Selected_for_final =
      (
        !is.na(
          table_now$
            `P value`
        ) &
          table_now$
          `P value` <
          cutoff
      ),
    stringsAsFactors = FALSE
  )
  
  
  result
}


preoperative_full_variable_tests <- extract_variable_selection_12(
  preoperative_full_model_12,
  cutoff = 0.05
)


precatheter_full_variable_tests <- extract_variable_selection_12(
  precatheter_full_model_12,
  cutoff = 0.05
)


# ----------------------------------------------------------
# 12.8 确定最终预测因子
# ----------------------------------------------------------

preoperative_final_vars <- (
  preoperative_full_variable_tests$
    Variable[
      preoperative_full_variable_tests$
        Selected_for_final
    ]
)


precatheter_final_vars <- (
  precatheter_full_variable_tests$
    Variable[
      precatheter_full_variable_tests$
        Selected_for_final
    ]
)


cat(
  "\n术前完整模型 pooled 结果：\n"
)

print(
  preoperative_full_variable_tests
)


cat(
  "\n术前最终预测因子（P < 0.05）：\n"
)

print(
  preoperative_final_vars
)


cat(
  "\n拔管前完整模型 pooled 结果：\n"
)

print(
  precatheter_full_variable_tests
)


cat(
  "\n拔管前最终预测因子（P < 0.05）：\n"
)

print(
  precatheter_final_vars
)


# ----------------------------------------------------------
# 12.8A 最终变量质控
# ----------------------------------------------------------

if (
  !all(
    preoperative_final_vars %in%
    preoperative_full_vars
  )
) {
  stop(
    "术前最终变量并非完整模型候选变量的子集。"
  )
}


if (
  !all(
    precatheter_final_vars %in%
    precatheter_full_vars
  )
) {
  stop(
    "拔管前最终变量并非完整模型候选变量的子集。"
  )
}


if (
  length(
    preoperative_final_vars
  ) == 0
) {
  
  warning(
    "术前完整模型中没有变量达到 pooled P < 0.05，将建立截距模型。"
  )
}


if (
  length(
    precatheter_final_vars
  ) == 0
) {
  
  warning(
    "拔管前完整模型中没有变量达到 pooled P < 0.05，将建立截距模型。"
  )
}


# ----------------------------------------------------------
# 12.9 保存完整模型结果和变量筛选过程
# ----------------------------------------------------------

openxlsx::write.xlsx(
  list(
    Preoperative =
      preoperative_full_model_12$
      pooled_table,
    Precatheter =
      precatheter_full_model_12$
      pooled_table
  ),
  file.path(
    table_dir,
    "Full_multivariable_models_pooled_results.xlsx"
  ),
  overwrite = TRUE
)


openxlsx::write.xlsx(
  list(
    Preoperative =
      preoperative_full_variable_tests,
    Precatheter =
      precatheter_full_variable_tests
  ),
  file.path(
    table_dir,
    "Full_model_variable_selection_P_lt_0.05.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 12.10 重新拟合两个最终模型
# ----------------------------------------------------------

preoperative_final_model_12 <- fit_mi_logistic_12(
  imp_object = imp_preoperative,
  predictor_vars = preoperative_final_vars,
  model_name = "Preoperative_final"
)


precatheter_final_model_12 <- fit_mi_logistic_12(
  imp_object = imp_precatheter,
  predictor_vars = precatheter_final_vars,
  model_name = "Precatheter_final"
)


# ----------------------------------------------------------
# 12.10A 最终模型质控
# ----------------------------------------------------------

final_model_qc <- dplyr::bind_rows(
  preoperative_final_model_12$
    qc_summary,
  precatheter_final_model_12$
    qc_summary
)


cat(
  "\n最终模型质控：\n"
)

print(
  final_model_qc
)


if (
  any(
    final_model_qc$
    Number_of_fits != 20
  )
) {
  stop(
    "最终模型没有全部完成20次拟合。"
  )
}


if (
  any(
    !final_model_qc$
    All_converged
  )
) {
  warning(
    "至少一个最终模型存在未收敛拟合。"
  )
}


if (
  any(
    final_model_qc$
    Any_extreme_coefficient
  )
) {
  warning(
    "至少一个最终模型出现|Beta|>10，请检查分离。"
  )
}


if (
  any(
    final_model_qc$
    Any_very_large_SE
  )
) {
  warning(
    "至少一个最终模型出现SE>5，请检查稀疏数据。"
  )
}


if (
  any(
    final_model_qc$
    Maximum_VIF_equivalent >= 5,
    na.rm = TRUE
  )
) {
  warning(
    "最终模型存在VIF equivalent >=5，请检查共线性。"
  )
}


# ----------------------------------------------------------
# 12.11 生成完整/最终多因素模型输出
#
# 输出表展示完整多因素模型中的全部候选变量，
# 即所有由Table 1中P < 0.10进入多因素分析的变量。
#
# 同时标记：
# - 是否满足完整多因素模型 pooled P < 0.05
# - 是否进入最终精简模型
#
# 对最终保留变量，再额外给出重新拟合最终模型后的参数。
#
# 这样输出表既展示完整排除过程，
# 又保留最终模型真正用于Nomogram和风险计算的系数。
# ----------------------------------------------------------

make_table2_complete_12 <- function(
    full_model_object,
    variable_selection_table,
    final_model_object,
    timepoint
) {
  
  full_table <- full_model_object$pooled_table
  
  full_table <- full_table[
    full_table$Variable != "(Intercept)",
    ,
    drop = FALSE
  ]
  
  full_result <- data.frame(
    Timepoint = timepoint,
    Variable = full_table$Variable,
    Term = full_table$Term,
    
    Full_Beta = round(
      full_table$Beta,
      3
    ),
    
    Full_SE = round(
      full_table$SE,
      3
    ),
    
    Full_Wald = round(
      full_table$Wald,
      3
    ),
    
    Full_P_value = full_table$
      `P formatted`,
    
    Full_OR_95CI = full_table$
      `OR (95% CI)`,
    
    stringsAsFactors = FALSE
  )
  
  
  # 加入是否进入最终模型
  selection_info <- variable_selection_table[
    ,
    c(
      "Variable",
      "Selected_for_final"
    ),
    drop = FALSE
  ]
  
  full_result <- dplyr::left_join(
    full_result,
    selection_info,
    by = "Variable"
  )
  
  if (any(is.na(full_result$Selected_for_final))) {
    stop(
      paste0(
        timepoint,
        " 多因素模型输出中存在无法匹配的变量筛选状态。"
      )
    )
  }
  
  full_result$Selected_for_final <- ifelse(
    full_result$Selected_for_final,
    "Yes",
    "No"
  )
  
  
  # 最终精简模型结果
  final_table <- final_model_object$pooled_table
  
  final_table <- final_table[
    final_table$Variable != "(Intercept)",
    ,
    drop = FALSE
  ]
  
  
  if (nrow(final_table) > 0) {
    
    final_result <- data.frame(
      Variable = final_table$Variable,
      
      Final_Beta = round(
        final_table$Beta,
        3
      ),
      
      Final_SE = round(
        final_table$SE,
        3
      ),
      
      Final_Wald = round(
        final_table$Wald,
        3
      ),
      
      Final_P_value = final_table$
        `P formatted`,
      
      Final_OR_95CI = final_table$
        `OR (95% CI)`,
      
      stringsAsFactors = FALSE
    )
    
    full_result <- dplyr::left_join(
      full_result,
      final_result,
      by = "Variable"
    )
    
  } else {
    
    full_result$Final_Beta <- NA_real_
    full_result$Final_SE <- NA_real_
    full_result$Final_Wald <- NA_real_
    full_result$Final_P_value <- NA_character_
    full_result$Final_OR_95CI <- NA_character_
  }
  
  
  full_result$Selection_reason <- ifelse(
    full_result$Selected_for_final == "Yes",
    "Retained: pooled multivariable P < 0.05",
    "Not retained: pooled multivariable P >= 0.05"
  )
  
  
  full_result <- full_result[
    ,
    c(
      "Timepoint",
      "Variable",
      "Term",
      
      "Full_Beta",
      "Full_SE",
      "Full_Wald",
      "Full_P_value",
      "Full_OR_95CI",
      
      "Selected_for_final",
      
      "Final_Beta",
      "Final_SE",
      "Final_Wald",
      "Final_P_value",
      "Final_OR_95CI",
      
      "Selection_reason"
    )
  ]
  
  
  return(full_result)
}


table2_preoperative <- make_table2_complete_12(
  full_model_object =
    preoperative_full_model_12,
  
  variable_selection_table =
    preoperative_full_variable_tests,
  
  final_model_object =
    preoperative_final_model_12,
  
  timepoint =
    "Preoperative"
)


table2_precatheter <- make_table2_complete_12(
  full_model_object =
    precatheter_full_model_12,
  
  variable_selection_table =
    precatheter_full_variable_tests,
  
  final_model_object =
    precatheter_final_model_12,
  
  timepoint =
    "Precatheter"
)


# ----------------------------------------------------------
# 12.11A Table 2质控
# ----------------------------------------------------------

if (
  nrow(table2_preoperative) !=
  length(preoperative_full_vars)
) {
  
  stop(
    "术前多因素模型输出行数与完整模型候选变量数不一致。"
  )
}


if (
  nrow(table2_precatheter) !=
  length(precatheter_full_vars)
) {
  
  stop(
    "拔管前多因素模型输出行数与完整模型候选变量数不一致。"
  )
}


if (
  sum(
    table2_preoperative$
    Selected_for_final ==
    "Yes"
  ) !=
  length(
    preoperative_final_vars
  )
) {
  
  stop(
    "术前多因素模型输出最终保留变量数量与最终模型不一致。"
  )
}


if (
  sum(
    table2_precatheter$
    Selected_for_final ==
    "Yes"
  ) !=
  length(
    precatheter_final_vars
  )
) {
  
  stop(
    "拔管前多因素模型输出最终保留变量数量与最终模型不一致。"
  )
}


if (
  any(
    is.na(
      table2_preoperative$
      Final_Beta[
        table2_preoperative$
        Selected_for_final ==
        "Yes"
      ]
    )
  )
) {
  
  stop(
    "术前多因素模型输出中存在最终保留变量但缺少最终模型参数。"
  )
}


if (
  any(
    is.na(
      table2_precatheter$
      Final_Beta[
        table2_precatheter$
        Selected_for_final ==
        "Yes"
      ]
    )
  )
) {
  
  stop(
    "拔管前多因素模型输出中存在最终保留变量但缺少最终模型参数。"
  )
}


cat(
  "12.11 多因素模型完整筛选过程质控：通过。\n"
)


# ----------------------------------------------------------
# 12.12 保存完整/最终多因素模型输出
#
# Preoperative：全部8个完整模型候选变量
# Precatheter：全部11个完整模型候选变量
#
# Full_*列：
#   完整多因素模型中的结果，用于展示筛选过程。
#
# Final_*列：
#   仅最终保留变量有结果，用于展示最终重拟合参数。
# ----------------------------------------------------------

table2_file <- file.path(
  table_dir,
  "Multivariable_logistic_models_full_and_final.xlsx"
)


openxlsx::write.xlsx(
  list(
    Preoperative =
      table2_preoperative,
    
    Precatheter =
      table2_precatheter
  ),
  table2_file,
  overwrite = TRUE
)


# ----------------------------------------------------------
# 12.12A 保存最终模型完整参数（含截距）
#
# 专供后续Nomogram、风险公式和网页计算器使用。
# 该参数文件供Nomogram、风险公式和网页计算器使用。
# ----------------------------------------------------------

final_model_parameter_table <- dplyr::bind_rows(
  
  data.frame(
    Timepoint = "Preoperative",
    preoperative_final_model_12$
      pooled_table,
    check.names = FALSE
  ),
  
  data.frame(
    Timepoint = "Precatheter",
    precatheter_final_model_12$
      pooled_table,
    check.names = FALSE
  )
)


openxlsx::write.xlsx(
  final_model_parameter_table,
  file.path(
    table_dir,
    "Final_model_pooled_coefficients_with_intercept.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 12.13 保存完整和最终模型VIF
# ----------------------------------------------------------

openxlsx::write.xlsx(
  list(
    Preoperative_full =
      preoperative_full_model_12$
      vif_detail,
    
    Precatheter_full =
      precatheter_full_model_12$
      vif_detail,
    
    Preoperative_final =
      preoperative_final_model_12$
      vif_detail,
    
    Precatheter_final =
      precatheter_final_model_12$
      vif_detail
  ),
  file.path(
    table_dir,
    "Multivariable_model_VIF_details.xlsx"
  ),
  overwrite = TRUE
)


openxlsx::write.xlsx(
  list(
    Preoperative_full =
      preoperative_full_model_12$
      vif_summary,
    
    Precatheter_full =
      precatheter_full_model_12$
      vif_summary,
    
    Preoperative_final =
      preoperative_final_model_12$
      vif_summary,
    
    Precatheter_final =
      precatheter_final_model_12$
      vif_summary
  ),
  file.path(
    table_dir,
    "Multivariable_model_VIF_summary.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 12.14 保存完整和最终模型质控
# ----------------------------------------------------------

all_model_qc <- dplyr::bind_rows(
  
  preoperative_full_model_12$
    qc_summary,
  
  precatheter_full_model_12$
    qc_summary,
  
  preoperative_final_model_12$
    qc_summary,
  
  precatheter_final_model_12$
    qc_summary
)


all_model_qc_detail <- dplyr::bind_rows(
  
  preoperative_full_model_12$
    qc_detail,
  
  precatheter_full_model_12$
    qc_detail,
  
  preoperative_final_model_12$
    qc_detail,
  
  precatheter_final_model_12$
    qc_detail
)


openxlsx::write.xlsx(
  list(
    Summary =
      all_model_qc,
    
    Detail =
      all_model_qc_detail
  ),
  file.path(
    log_dir,
    "Multivariable_model_QC.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 12.15 保存最终模型公式 + 表观模型性能
#
# 注意：
# AUC、Brier、IPA、Log loss均为20个MI数据集上的
# apparent performance描述性汇总（mean/SD across MI）。
# 不能替代后续bootstrap/CV内部验证结果。
# ----------------------------------------------------------

final_model_formula_table <- data.frame(
  Model = c(
    "Preoperative",
    "Precatheter"
  ),
  
  Formula = c(
    paste(
      deparse(
        preoperative_final_model_12$formula
      ),
      collapse = " "
    ),
    paste(
      deparse(
        precatheter_final_model_12$formula
      ),
      collapse = " "
    )
  ),
  
  Apparent_AUC = round(
    c(
      preoperative_final_model_12$qc_summary$Mean_Apparent_AUC,
      precatheter_final_model_12$qc_summary$Mean_Apparent_AUC
    ),
    3
  ),
  
  AUC_SD_across_MI = round(
    c(
      preoperative_final_model_12$qc_summary$SD_Apparent_AUC,
      precatheter_final_model_12$qc_summary$SD_Apparent_AUC
    ),
    3
  ),
  
  Brier_Score = round(
    c(
      preoperative_final_model_12$qc_summary$Mean_Brier,
      precatheter_final_model_12$qc_summary$Mean_Brier
    ),
    4
  ),
  
  Brier_SD_across_MI = round(
    c(
      preoperative_final_model_12$qc_summary$SD_Brier,
      precatheter_final_model_12$qc_summary$SD_Brier
    ),
    4
  ),
  
  Scaled_Brier_IPA = round(
    c(
      preoperative_final_model_12$qc_summary$Mean_Scaled_Brier_IPA,
      precatheter_final_model_12$qc_summary$Mean_Scaled_Brier_IPA
    ),
    4
  ),
  
  Log_Loss = round(
    c(
      preoperative_final_model_12$qc_summary$Mean_Log_Loss,
      precatheter_final_model_12$qc_summary$Mean_Log_Loss
    ),
    4
  ),
  
  Mean_AIC = round(
    c(
      preoperative_final_model_12$qc_summary$Mean_AIC,
      precatheter_final_model_12$qc_summary$Mean_AIC
    ),
    2
  ),
  
  Mean_BIC = round(
    c(
      preoperative_final_model_12$qc_summary$Mean_BIC,
      precatheter_final_model_12$qc_summary$Mean_BIC
    ),
    2
  ),
  
  Observed_Event_Rate = round(
    c(
      preoperative_final_model_12$qc_summary$Mean_Observed_Event_Rate,
      precatheter_final_model_12$qc_summary$Mean_Observed_Event_Rate
    ),
    4
  ),
  
  Mean_Predicted_Risk = round(
    c(
      preoperative_final_model_12$qc_summary$Mean_Predicted_Risk,
      precatheter_final_model_12$qc_summary$Mean_Predicted_Risk
    ),
    4
  ),
  
  Performance_type = c(
    "Apparent performance; mean across 20 MI datasets; not internally validated",
    "Apparent performance; mean across 20 MI datasets; not internally validated"
  ),
  
  stringsAsFactors = FALSE
)


openxlsx::write.xlsx(
  final_model_formula_table,
  file.path(
    table_dir,
    "Final_model_formulas.xlsx"
  ),
  overwrite = TRUE
)


# ----------------------------------------------------------
# 12.16 保存模型对象
# ----------------------------------------------------------

saveRDS(
  list(
    Preoperative_full =
      preoperative_full_model_12,
    
    Precatheter_full =
      precatheter_full_model_12,
    
    Preoperative_final =
      preoperative_final_model_12,
    
    Precatheter_final =
      precatheter_final_model_12
  ),
  file.path(
    model_dir,
    "Final_multivariable_model_objects.rds"
  )
)


# ----------------------------------------------------------
# 12.17 最终文件质控
# ----------------------------------------------------------

required_output_files_12 <- c(
  
  file.path(
    table_dir,
    "Final_route_full_model_candidate_variables.xlsx"
  ),
  
  file.path(
    table_dir,
    "Full_multivariable_models_pooled_results.xlsx"
  ),
  
  file.path(
    table_dir,
    "Full_model_variable_selection_P_lt_0.05.xlsx"
  ),
  
  file.path(
    table_dir,
    "Multivariable_logistic_models_full_and_final.xlsx"
  ),
  
  file.path(
    table_dir,
    "Final_model_pooled_coefficients_with_intercept.xlsx"
  ),
  
  file.path(
    table_dir,
    "Multivariable_model_VIF_summary.xlsx"
  ),
  
  file.path(
    log_dir,
    "Multivariable_model_QC.xlsx"
  ),
  
  file.path(
    table_dir,
    "Final_model_formulas.xlsx"
  ),
  
  file.path(
    model_dir,
    "Final_multivariable_model_objects.rds"
  )
)


missing_output_files_12 <- required_output_files_12[
  !file.exists(
    required_output_files_12
  )
]


if (length(missing_output_files_12) > 0) {
  
  stop(
    paste0(
      "第12部分以下输出文件未成功生成：",
      paste(
        basename(
          missing_output_files_12
        ),
        collapse = ", "
      )
    )
  )
}


cat(
  "12.17 最终输出文件质控：通过。\n"
)


# ----------------------------------------------------------
# 12.18 控制台输出最终结果
# ----------------------------------------------------------

cat(
  "\n术前多因素模型输出（完整筛选过程）：\n"
)

print(
  table2_preoperative
)


cat(
  "\n拔管前多因素模型输出（完整筛选过程）：\n"
)

print(
  table2_precatheter
)


cat(
  "\n最终模型公式：\n"
)

print(
  final_model_formula_table
)


cat(
  "\n完整/最终模型质控：\n"
)

print(
  all_model_qc
)


cat(
  "\n>>> 第12部分全部完成 <<<\n"
)


cat(
  "请重点检查：\n",
  "02_tables/Multivariable_logistic_models_full_and_final.xlsx\n",
  "02_tables/Final_model_pooled_coefficients_with_intercept.xlsx\n",
  "02_tables/Multivariable_model_VIF_summary.xlsx\n",
  "05_logs/Multivariable_model_QC.xlsx\n",
  "02_tables/Final_model_formulas.xlsx\n"
)



############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
# 13. 最终模型 Bootstrap 内部验证（1000次，Calibration修正版）
#
# 目的：
# 1. 对第12部分已经锁定的两个最终 Logistic 模型进行内部验证；
# 2. 使用 ordinary nonparametric bootstrap，B = 1000；
# 3. 两个时间点均使用相同的bootstrap抽样种子，便于结果比较；
# 4. 在20个MICE完整数据集中使用相同的bootstrap患者索引；
# 5. 每次bootstrap均重新拟合“已经锁定的最终模型公式”；
# 6. 计算：
#       - AUC
#       - Brier score
#       - Scaled Brier / IPA
#       - Log loss
#       - Calibration intercept
#       - Calibration slope
# 7. 对每个指标计算：
#       Apparent performance
#       Mean optimism
#       Optimism-corrected performance
# 8. 本节验证的是“锁定后的最终模型”，
#    不在bootstrap内部重新执行Table 1筛选或变量选择。
#
# 重要：
# - 本节采用现有20个MI数据集进行MI-aware bootstrap；
# - 不在每个bootstrap样本内重新运行MICE；
# - 对当前缺失比例较低的数据，这是实用且可复核的内部验证方案；
# - Calibration intercept采用稳定的score-root方法，避免极端bootstrap样本导致glm伪发散；
# - 后续Nomogram应以第12部分Rubin pooled coefficients为基础，
#   本节主要用于评价模型的过拟合与内部性能。
############################################################


cat("\n>>> 第13部分开始执行：Bootstrap内部验证 <<<\n")
flush.console()


# ==========================================================
# 13.1 参数设置
# ==========================================================

B_13 <- 1000L
seed_13 <- 20260817L

# 每多少次bootstrap输出一次进度
progress_every_13 <- 100L

# 概率裁剪，避免qlogis/log出现Inf
epsilon_13 <- 1e-12


# ==========================================================
# 13.2 运行前检查
# ==========================================================

required_objects_13 <- c(
  "imp_preoperative",
  "imp_precatheter",
  "preoperative_final_model_12",
  "precatheter_final_model_12",
  "outcome_var",
  "table_dir",
  "model_dir",
  "log_dir"
)

missing_objects_13 <- required_objects_13[
  !vapply(
    required_objects_13,
    exists,
    logical(1),
    inherits = TRUE
  )
]

if (length(missing_objects_13) > 0) {
  stop(
    paste0(
      "第13部分缺少前置对象：",
      paste(
        missing_objects_13,
        collapse = ", "
      )
    )
  )
}

required_packages_13 <- c(
  "mice",
  "dplyr",
  "openxlsx"
)

missing_packages_13 <- required_packages_13[
  !vapply(
    required_packages_13,
    requireNamespace,
    logical(1),
    quietly = TRUE
  )
]

if (length(missing_packages_13) > 0) {
  stop(
    paste0(
      "第13部分缺少R包：",
      paste(
        missing_packages_13,
        collapse = ", "
      )
    )
  )
}

if (imp_preoperative$m != 20) {
  stop("术前MICE对象不是20个插补数据集。")
}

if (imp_precatheter$m != 20) {
  stop("拔管前MICE对象不是20个插补数据集。")
}

if (B_13 < 200) {
  warning("Bootstrap次数少于200；正式结果建议至少1000次。")
}

dir.create(
  table_dir,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  model_dir,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  log_dir,
  recursive = TRUE,
  showWarnings = FALSE
)

cat(
  "13.2 运行前检查：通过。\n",
  "Bootstrap次数 = ",
  B_13,
  "\n",
  "随机种子 = ",
  seed_13,
  "\n",
  sep = ""
)


# ==========================================================
# 13.3 基础辅助函数
# ==========================================================

# ----------------------------------------------------------
# 13.3A 将结局安全转换为0/1
# ----------------------------------------------------------

coerce_binary_outcome_13 <- function(
    x,
    variable_name = "Outcome"
) {
  
  if (is.logical(x)) {
    x <- as.integer(x)
  }
  
  if (is.factor(x) || is.character(x)) {
    
    x_character <- trimws(
      as.character(x)
    )
    
    valid_values <- unique(
      x_character[
        !is.na(x_character)
      ]
    )
    
    if (
      !all(
        valid_values %in%
        c("0", "1")
      )
    ) {
      stop(
        paste0(
          variable_name,
          " 不是明确的0/1编码。当前取值为：",
          paste(
            valid_values,
            collapse = ", "
          )
        )
      )
    }
    
    x <- as.numeric(
      x_character
    )
  }
  
  x <- as.numeric(x)
  
  if (anyNA(x)) {
    stop(
      paste0(
        variable_name,
        " 转换为0/1后出现NA。"
      )
    )
  }
  
  if (!all(x %in% c(0, 1))) {
    stop(
      paste0(
        variable_name,
        " 不是0/1二分类结局。"
      )
    )
  }
  
  x
}


# ----------------------------------------------------------
# 13.3B 快速AUC
#
# 与Mann-Whitney / rank-sum定义等价，
# 避免1000×20×2次反复调用pROC，提高运行速度。
# ----------------------------------------------------------

fast_auc_13 <- function(
    y,
    p
) {
  
  ok <- is.finite(y) &
    is.finite(p)
  
  y <- y[ok]
  p <- p[ok]
  
  n1 <- sum(y == 1)
  n0 <- sum(y == 0)
  
  if (
    n1 == 0 ||
    n0 == 0
  ) {
    return(NA_real_)
  }
  
  r <- rank(
    p,
    ties.method = "average"
  )
  
  auc <- (
    sum(
      r[y == 1]
    ) -
      n1 *
      (n1 + 1) /
      2
  ) /
    (
      n1 *
        n0
    )
  
  as.numeric(auc)
}


# ----------------------------------------------------------
# 13.3C 基础预测性能
# ----------------------------------------------------------

basic_performance_13 <- function(
    y,
    p
) {
  
  p <- pmin(
    pmax(
      as.numeric(p),
      epsilon_13
    ),
    1 - epsilon_13
  )
  
  y <- as.numeric(y)
  
  auc_value <- fast_auc_13(
    y = y,
    p = p
  )
  
  brier_value <- mean(
    (y - p)^2
  )
  
  event_rate <- mean(y)
  
  null_brier <- mean(
    (y - event_rate)^2
  )
  
  ipa_value <- if (
    is.finite(null_brier) &&
    null_brier > 0
  ) {
    1 -
      brier_value /
      null_brier
  } else {
    NA_real_
  }
  
  log_loss_value <- -mean(
    y * log(p) +
      (1 - y) *
      log(1 - p)
  )
  
  c(
    AUC = auc_value,
    Brier = brier_value,
    IPA = ipa_value,
    LogLoss = log_loss_value,
    EventRate = event_rate,
    MeanPredictedRisk = mean(p),
    NullBrier = null_brier
  )
}


# ----------------------------------------------------------
# 13.3D Calibration intercept / slope
#
# Calibration intercept:
#   固定预测线性项 lp = logit(p)，仅重新估计一个截距alpha：
#       logit(P(Y=1)) = alpha + lp
#
#   这里不再直接使用 glm(y ~ 1, offset = lp)，
#   因为在少数bootstrap样本中，极端预测概率可能导致
#   glm数值迭代不收敛并返回极大的伪截距。
#
#   改用score equation的单调根求解：
#       sum[y - plogis(alpha + lp)] = 0
#
#   该方法与“固定slope=1的calibration-in-the-large”
#   定义一致，但数值上更加稳定。
#
# Calibration slope:
#   glm(y ~ lp)
#   只有在模型真正收敛且系数有限时才接受；
#   否则返回NA，不让非收敛结果污染bootstrap均值。
#
# 理想值：
#   Calibration intercept = 0
#   Calibration slope     = 1
# ----------------------------------------------------------

calibration_metrics_13 <- function(
    y,
    p
) {
  
  y <- as.numeric(y)
  
  p <- pmin(
    pmax(
      as.numeric(p),
      epsilon_13
    ),
    1 - epsilon_13
  )
  
  ok <- is.finite(y) &
    is.finite(p)
  
  y <- y[ok]
  p <- p[ok]
  
  if (
    length(y) == 0 ||
    length(unique(y)) < 2
  ) {
    return(
      c(
        CalibrationIntercept = NA_real_,
        CalibrationSlope = NA_real_
      )
    )
  }
  
  lp <- stats::qlogis(p)
  
  if (
    any(
      !is.finite(lp)
    )
  ) {
    return(
      c(
        CalibrationIntercept = NA_real_,
        CalibrationSlope = NA_real_
      )
    )
  }
  
  
  # --------------------------------------------------------
  # A. Calibration intercept
  #
  # 求解：
  # sum(y - plogis(alpha + lp)) = 0
  #
  # 该score函数对alpha严格单调下降；
  # 只要结局同时包含0和1，理论上一定存在有限根。
  # --------------------------------------------------------
  
  intercept_score_13 <- function(alpha) {
    
    sum(
      y -
        stats::plogis(
          alpha + lp
        )
    )
  }
  
  
  lower_13 <- -10
  upper_13 <- 10
  
  f_lower_13 <- intercept_score_13(
    lower_13
  )
  
  f_upper_13 <- intercept_score_13(
    upper_13
  )
  
  
  # 若初始区间未包住根，逐步扩展。
  # 正常情况下几乎不会进入很多轮。
  expand_count_13 <- 0L
  
  while (
    is.finite(f_lower_13) &&
    is.finite(f_upper_13) &&
    f_lower_13 * f_upper_13 > 0 &&
    expand_count_13 < 20L
  ) {
    
    lower_13 <- lower_13 - 10
    upper_13 <- upper_13 + 10
    
    f_lower_13 <- intercept_score_13(
      lower_13
    )
    
    f_upper_13 <- intercept_score_13(
      upper_13
    )
    
    expand_count_13 <- expand_count_13 + 1L
  }
  
  
  calibration_intercept <- if (
    !is.finite(f_lower_13) ||
    !is.finite(f_upper_13) ||
    f_lower_13 * f_upper_13 > 0
  ) {
    
    NA_real_
    
  } else {
    
    intercept_root_13 <- tryCatch(
      
      stats::uniroot(
        f = intercept_score_13,
        interval = c(
          lower_13,
          upper_13
        ),
        tol = 1e-10,
        maxiter = 1000L
      ),
      
      error = function(e) NULL
    )
    
    
    if (
      is.null(
        intercept_root_13
      ) ||
      !is.finite(
        intercept_root_13$root
      )
    ) {
      
      NA_real_
      
    } else {
      
      as.numeric(
        intercept_root_13$root
      )
    }
  }
  
  
  # --------------------------------------------------------
  # B. Calibration slope
  #
  # 若lp几乎为常数，则无法估计slope。
  # 对非收敛或非有限系数直接返回NA。
  # --------------------------------------------------------
  
  if (
    length(
      unique(
        round(
          lp,
          12
        )
      )
    ) <= 1
  ) {
    
    calibration_slope <- NA_real_
    
  } else {
    
    slope_fit <- tryCatch(
      
      suppressWarnings(
        stats::glm(
          y ~ lp,
          family = stats::binomial(),
          control = stats::glm.control(
            epsilon = 1e-10,
            maxit = 100L
          )
        )
      ),
      
      error = function(e) NULL
    )
    
    
    calibration_slope <- if (
      is.null(
        slope_fit
      ) ||
      !isTRUE(
        slope_fit$converged
      ) ||
      length(
        stats::coef(
          slope_fit
        )
      ) < 2 ||
      !is.finite(
        stats::coef(
          slope_fit
        )[2]
      )
    ) {
      
      NA_real_
      
    } else {
      
      as.numeric(
        stats::coef(
          slope_fit
        )[2]
      )
    }
  }
  
  
  c(
    CalibrationIntercept =
      calibration_intercept,
    CalibrationSlope =
      calibration_slope
  )
}


# ----------------------------------------------------------
# 13.3E 安全均值 / SD / 分位数
# ----------------------------------------------------------

safe_mean_13 <- function(x) {
  
  if (
    length(x) == 0 ||
    all(is.na(x))
  ) {
    return(NA_real_)
  }
  
  mean(
    x,
    na.rm = TRUE
  )
}


safe_sd_13 <- function(x) {
  
  if (
    sum(
      !is.na(x)
    ) <= 1
  ) {
    return(NA_real_)
  }
  
  stats::sd(
    x,
    na.rm = TRUE
  )
}


safe_quantile_13 <- function(
    x,
    prob
) {
  
  if (
    length(x) == 0 ||
    all(is.na(x))
  ) {
    return(NA_real_)
  }
  
  as.numeric(
    stats::quantile(
      x,
      probs = prob,
      na.rm = TRUE,
      names = FALSE,
      type = 7
    )
  )
}


# ==========================================================
# 13.4 单个最终模型的MI-aware Bootstrap内部验证
# ==========================================================

bootstrap_validate_mi_model_13 <- function(
    imp_object,
    model_object,
    model_label,
    B = 1000L,
    seed = 20260817L,
    progress_every = 100L
) {
  
  m <- imp_object$m
  
  if (
    is.null(
      model_object$formula
    )
  ) {
    stop(
      paste0(
        model_label,
        " 缺少formula。"
      )
    )
  }
  
  if (
    is.null(
      model_object$fits
    ) ||
    length(
      model_object$fits
    ) != m
  ) {
    stop(
      paste0(
        model_label,
        " 的最终模型对象中没有与MICE数量一致的fits。"
      )
    )
  }
  
  formula_object <- model_object$formula
  
  predictor_vars <- attr(
    stats::terms(
      formula_object
    ),
    "term.labels"
  )
  
  cat(
    "\n====================================================\n",
    "开始Bootstrap内部验证：",
    model_label,
    "\n公式：",
    paste(
      deparse(
        formula_object
      ),
      collapse = " "
    ),
    "\n预测变量数：",
    length(
      predictor_vars
    ),
    "\n====================================================\n",
    sep = ""
  )
  
  # --------------------------------------------------------
  # 13.4A 预先读取20个完整插补数据集
  # --------------------------------------------------------
  
  completed_list <- vector(
    "list",
    m
  )
  
  required_vars <- c(
    outcome_var,
    predictor_vars
  )
  
  n_vector <- integer(m)
  
  for (i in seq_len(m)) {
    
    data_i <- mice::complete(
      imp_object,
      action = i
    )
    
    missing_vars <- setdiff(
      required_vars,
      names(data_i)
    )
    
    if (
      length(
        missing_vars
      ) > 0
    ) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个插补数据集缺少变量：",
          paste(
            missing_vars,
            collapse = ", "
          )
        )
      )
    }
    
    if (
      anyNA(
        data_i[
          required_vars
        ]
      )
    ) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个插补数据集的最终模型变量仍有缺失。"
        )
      )
    }
    
    completed_list[[i]] <- data_i
    n_vector[i] <- nrow(data_i)
  }
  
  if (
    length(
      unique(
        n_vector
      )
    ) != 1
  ) {
    stop(
      paste0(
        model_label,
        " 的20个插补数据集样本量不一致。"
      )
    )
  }
  
  n <- n_vector[1]
  
  y_reference <- coerce_binary_outcome_13(
    completed_list[[1]][[outcome_var]],
    variable_name = outcome_var
  )
  
  # 确保20个插补数据集的结局完全一致且行顺序一致
  for (i in seq_len(m)) {
    
    y_i <- coerce_binary_outcome_13(
      completed_list[[i]][[outcome_var]],
      variable_name = outcome_var
    )
    
    if (
      !identical(
        y_reference,
        y_i
      )
    ) {
      stop(
        paste0(
          model_label,
          " 的20个插补数据集结局顺序不一致。"
        )
      )
    }
  }
  
  event_n <- sum(
    y_reference == 1
  )
  
  cat(
    "样本量 N = ",
    n,
    "\n事件数 = ",
    event_n,
    "\nMICE数据集 = ",
    m,
    "\n",
    sep = ""
  )
  
  # --------------------------------------------------------
  # 13.4B 计算表观性能
  #
  # AUC/Brier/LogLoss：
  #   在每个MI数据集的原始最终模型上计算，再取均值。
  #
  # Calibration：
  #   对20个MI模型预测概率逐患者取平均后评价。
  # --------------------------------------------------------
  
  apparent_by_mi_list <- vector(
    "list",
    m
  )
  
  apparent_prediction_matrix <- matrix(
    NA_real_,
    nrow = n,
    ncol = m
  )
  
  for (i in seq_len(m)) {
    
    fit_i <- model_object$fits[[i]]
    data_i <- completed_list[[i]]
    
    prediction_i <- tryCatch(
      stats::predict(
        fit_i,
        newdata = data_i,
        type = "response"
      ),
      error = function(e) NULL
    )
    
    if (
      is.null(
        prediction_i
      ) ||
      length(
        prediction_i
      ) != n ||
      any(
        !is.finite(
          prediction_i
        )
      )
    ) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个原始最终模型无法生成有效预测概率。"
        )
      )
    }
    
    apparent_prediction_matrix[, i] <- prediction_i
    
    perf_i <- basic_performance_13(
      y = y_reference,
      p = prediction_i
    )
    
    apparent_by_mi_list[[i]] <- data.frame(
      Model = model_label,
      Imputation = i,
      AUC = unname(
        perf_i["AUC"]
      ),
      Brier = unname(
        perf_i["Brier"]
      ),
      IPA = unname(
        perf_i["IPA"]
      ),
      LogLoss = unname(
        perf_i["LogLoss"]
      ),
      EventRate = unname(
        perf_i["EventRate"]
      ),
      MeanPredictedRisk = unname(
        perf_i["MeanPredictedRisk"]
      ),
      stringsAsFactors = FALSE
    )
  }
  
  apparent_by_mi <- dplyr::bind_rows(
    apparent_by_mi_list
  )
  
  apparent_prediction_mean <- rowMeans(
    apparent_prediction_matrix
  )
  
  apparent_calibration <- calibration_metrics_13(
    y = y_reference,
    p = apparent_prediction_mean
  )
  
  apparent_auc <- safe_mean_13(
    apparent_by_mi$AUC
  )
  
  apparent_brier <- safe_mean_13(
    apparent_by_mi$Brier
  )
  
  apparent_logloss <- safe_mean_13(
    apparent_by_mi$LogLoss
  )
  
  original_event_rate <- mean(
    y_reference
  )
  
  original_null_brier <- mean(
    (
      y_reference -
        original_event_rate
    )^2
  )
  
  apparent_ipa <- if (
    original_null_brier > 0
  ) {
    1 -
      apparent_brier /
      original_null_brier
  } else {
    NA_real_
  }
  
  apparent_calibration_intercept <- unname(
    apparent_calibration[
      "CalibrationIntercept"
    ]
  )
  
  apparent_calibration_slope <- unname(
    apparent_calibration[
      "CalibrationSlope"
    ]
  )
  
  cat(
    "\n表观性能：\n",
    "AUC = ",
    sprintf(
      "%.4f",
      apparent_auc
    ),
    "\nBrier = ",
    sprintf(
      "%.4f",
      apparent_brier
    ),
    "\nIPA = ",
    sprintf(
      "%.4f",
      apparent_ipa
    ),
    "\nCalibration intercept = ",
    sprintf(
      "%.4f",
      apparent_calibration_intercept
    ),
    "\nCalibration slope = ",
    sprintf(
      "%.4f",
      apparent_calibration_slope
    ),
    "\n",
    sep = ""
  )
  
  # --------------------------------------------------------
  # 13.4C 生成bootstrap索引
  #
  # 对20个MI数据集使用完全相同的患者索引。
  # --------------------------------------------------------
  
  set.seed(seed)
  
  bootstrap_indices <- replicate(
    B,
    sample.int(
      n = n,
      size = n,
      replace = TRUE
    ),
    simplify = FALSE
  )
  
  # --------------------------------------------------------
  # 13.4D Bootstrap循环
  # --------------------------------------------------------
  
  bootstrap_detail_list <- vector(
    "list",
    B
  )
  
  failed_bootstrap_ids <- integer(0)
  
  total_fit_warnings <- 0L
  
  for (b in seq_len(B)) {
    
    if (
      b == 1L ||
      b %% progress_every == 0L ||
      b == B
    ) {
      cat(
        "  Bootstrap ",
        b,
        "/",
        B,
        "\n",
        sep = ""
      )
      flush.console()
    }
    
    index_b <- bootstrap_indices[[b]]
    
    y_boot <- y_reference[
      index_b
    ]
    
    # 极少数情况下bootstrap样本若只有一个结局水平则无法验证
    if (
      length(
        unique(
          y_boot
        )
      ) < 2
    ) {
      
      failed_bootstrap_ids <- c(
        failed_bootstrap_ids,
        b
      )
      
      next
    }
    
    train_prediction_matrix <- matrix(
      NA_real_,
      nrow = n,
      ncol = m
    )
    
    test_prediction_matrix <- matrix(
      NA_real_,
      nrow = n,
      ncol = m
    )
    
    train_auc_mi <- rep(
      NA_real_,
      m
    )
    
    test_auc_mi <- rep(
      NA_real_,
      m
    )
    
    train_brier_mi <- rep(
      NA_real_,
      m
    )
    
    test_brier_mi <- rep(
      NA_real_,
      m
    )
    
    train_logloss_mi <- rep(
      NA_real_,
      m
    )
    
    test_logloss_mi <- rep(
      NA_real_,
      m
    )
    
    bootstrap_failed <- FALSE
    
    for (i in seq_len(m)) {
      
      data_i <- completed_list[[i]]
      
      data_boot <- data_i[
        index_b,
        ,
        drop = FALSE
      ]
      
      captured_warnings <- character()
      
      fit_boot <- tryCatch(
        
        withCallingHandlers(
          
          stats::glm(
            formula = formula_object,
            data = data_boot,
            family = stats::binomial()
          ),
          
          warning = function(w) {
            
            captured_warnings <<- c(
              captured_warnings,
              conditionMessage(w)
            )
            
            invokeRestart(
              "muffleWarning"
            )
          }
        ),
        
        error = function(e) NULL
      )
      
      total_fit_warnings <- total_fit_warnings +
        length(
          captured_warnings
        )
      
      if (
        is.null(
          fit_boot
        ) ||
        !isTRUE(
          fit_boot$converged
        ) ||
        any(
          !is.finite(
            stats::coef(
              fit_boot
            )
          )
        )
      ) {
        
        bootstrap_failed <- TRUE
        break
      }
      
      pred_train_i <- tryCatch(
        stats::predict(
          fit_boot,
          newdata = data_boot,
          type = "response"
        ),
        error = function(e) NULL
      )
      
      pred_test_i <- tryCatch(
        stats::predict(
          fit_boot,
          newdata = data_i,
          type = "response"
        ),
        error = function(e) NULL
      )
      
      if (
        is.null(
          pred_train_i
        ) ||
        is.null(
          pred_test_i
        ) ||
        length(
          pred_train_i
        ) != n ||
        length(
          pred_test_i
        ) != n ||
        any(
          !is.finite(
            pred_train_i
          )
        ) ||
        any(
          !is.finite(
            pred_test_i
          )
        )
      ) {
        
        bootstrap_failed <- TRUE
        break
      }
      
      train_prediction_matrix[, i] <-
        pred_train_i
      
      test_prediction_matrix[, i] <-
        pred_test_i
      
      train_perf_i <- basic_performance_13(
        y = y_boot,
        p = pred_train_i
      )
      
      test_perf_i <- basic_performance_13(
        y = y_reference,
        p = pred_test_i
      )
      
      train_auc_mi[i] <- unname(
        train_perf_i["AUC"]
      )
      
      test_auc_mi[i] <- unname(
        test_perf_i["AUC"]
      )
      
      train_brier_mi[i] <- unname(
        train_perf_i["Brier"]
      )
      
      test_brier_mi[i] <- unname(
        test_perf_i["Brier"]
      )
      
      train_logloss_mi[i] <- unname(
        train_perf_i["LogLoss"]
      )
      
      test_logloss_mi[i] <- unname(
        test_perf_i["LogLoss"]
      )
    }
    
    if (bootstrap_failed) {
      
      failed_bootstrap_ids <- c(
        failed_bootstrap_ids,
        b
      )
      
      next
    }
    
    # AUC/Brier/LogLoss按20个MI数据集取平均
    train_auc <- safe_mean_13(
      train_auc_mi
    )
    
    test_auc <- safe_mean_13(
      test_auc_mi
    )
    
    train_brier <- safe_mean_13(
      train_brier_mi
    )
    
    test_brier <- safe_mean_13(
      test_brier_mi
    )
    
    train_logloss <- safe_mean_13(
      train_logloss_mi
    )
    
    test_logloss <- safe_mean_13(
      test_logloss_mi
    )
    
    # Calibration使用MI平均预测概率
    train_prediction_mean <- rowMeans(
      train_prediction_matrix
    )
    
    test_prediction_mean <- rowMeans(
      test_prediction_matrix
    )
    
    train_calibration <- calibration_metrics_13(
      y = y_boot,
      p = train_prediction_mean
    )
    
    test_calibration <- calibration_metrics_13(
      y = y_reference,
      p = test_prediction_mean
    )
    
    train_calibration_intercept <- unname(
      train_calibration[
        "CalibrationIntercept"
      ]
    )
    
    test_calibration_intercept <- unname(
      test_calibration[
        "CalibrationIntercept"
      ]
    )
    
    train_calibration_slope <- unname(
      train_calibration[
        "CalibrationSlope"
      ]
    )
    
    test_calibration_slope <- unname(
      test_calibration[
        "CalibrationSlope"
      ]
    )
    
    # 统一定义：
    # optimism = bootstrap-training performance - original-test performance
    #
    # 对AUC：通常为正
    # 对Brier/LogLoss：通常为负
    # corrected = apparent - mean(optimism)
    bootstrap_detail_list[[b]] <- data.frame(
      
      Model = model_label,
      Bootstrap = b,
      
      Bootstrap_Events = sum(
        y_boot == 1
      ),
      
      Train_AUC = train_auc,
      Test_AUC = test_auc,
      Optimism_AUC =
        train_auc -
        test_auc,
      
      Train_Brier = train_brier,
      Test_Brier = test_brier,
      Optimism_Brier =
        train_brier -
        test_brier,
      
      Train_LogLoss = train_logloss,
      Test_LogLoss = test_logloss,
      Optimism_LogLoss =
        train_logloss -
        test_logloss,
      
      Train_Calibration_Intercept =
        train_calibration_intercept,
      
      Test_Calibration_Intercept =
        test_calibration_intercept,
      
      Optimism_Calibration_Intercept =
        train_calibration_intercept -
        test_calibration_intercept,
      
      Train_Calibration_Slope =
        train_calibration_slope,
      
      Test_Calibration_Slope =
        test_calibration_slope,
      
      Optimism_Calibration_Slope =
        train_calibration_slope -
        test_calibration_slope,
      
      stringsAsFactors = FALSE
    )
  }
  
  # --------------------------------------------------------
  # 13.4E 汇总有效bootstrap
  # --------------------------------------------------------
  
  valid_bootstrap_list <- bootstrap_detail_list[
    !vapply(
      bootstrap_detail_list,
      is.null,
      logical(1)
    )
  ]
  
  if (
    length(
      valid_bootstrap_list
    ) == 0
  ) {
    stop(
      paste0(
        model_label,
        " 没有任何有效bootstrap重复。"
      )
    )
  }
  
  bootstrap_detail <- dplyr::bind_rows(
    valid_bootstrap_list
  )
  
  valid_B <- nrow(
    bootstrap_detail
  )
  
  if (
    valid_B <
    0.95 *
    B
  ) {
    warning(
      paste0(
        model_label,
        " 仅有 ",
        valid_B,
        "/",
        B,
        " 次有效bootstrap；请检查稀疏数据或分离问题。"
      )
    )
  }
  
  # --------------------------------------------------------
  # 13.4F 计算平均optimism
  # --------------------------------------------------------
  
  mean_optimism_auc <- safe_mean_13(
    bootstrap_detail$
      Optimism_AUC
  )
  
  mean_optimism_brier <- safe_mean_13(
    bootstrap_detail$
      Optimism_Brier
  )
  
  mean_optimism_logloss <- safe_mean_13(
    bootstrap_detail$
      Optimism_LogLoss
  )
  
  mean_optimism_calibration_intercept <-
    safe_mean_13(
      bootstrap_detail$
        Optimism_Calibration_Intercept
    )
  
  mean_optimism_calibration_slope <-
    safe_mean_13(
      bootstrap_detail$
        Optimism_Calibration_Slope
    )
  
  # --------------------------------------------------------
  # 13.4G Optimism-corrected performance
  # --------------------------------------------------------
  
  corrected_auc <-
    apparent_auc -
    mean_optimism_auc
  
  corrected_brier <-
    apparent_brier -
    mean_optimism_brier
  
  corrected_logloss <-
    apparent_logloss -
    mean_optimism_logloss
  
  corrected_calibration_intercept <-
    apparent_calibration_intercept -
    mean_optimism_calibration_intercept
  
  corrected_calibration_slope <-
    apparent_calibration_slope -
    mean_optimism_calibration_slope
  
  corrected_ipa <- if (
    original_null_brier > 0
  ) {
    1 -
      corrected_brier /
      original_null_brier
  } else {
    NA_real_
  }
  
  # --------------------------------------------------------
  # 13.4H 构造每次bootstrap对应的corrected estimate
  #
  # 这些分位数用于描述bootstrap分布，
  # 不应机械解释为独立外部验证意义上的95%CI。
  # --------------------------------------------------------
  
  bootstrap_detail$Corrected_AUC <- apparent_auc -
    bootstrap_detail$
    Optimism_AUC
  
  bootstrap_detail$Corrected_Brier <- apparent_brier -
    bootstrap_detail$
    Optimism_Brier
  
  bootstrap_detail$Corrected_LogLoss <- apparent_logloss -
    bootstrap_detail$
    Optimism_LogLoss
  
  bootstrap_detail$Corrected_Calibration_Intercept <-
    apparent_calibration_intercept -
    bootstrap_detail$
    Optimism_Calibration_Intercept
  
  bootstrap_detail$Corrected_Calibration_Slope <-
    apparent_calibration_slope -
    bootstrap_detail$
    Optimism_Calibration_Slope
  
  # --------------------------------------------------------
  # 13.4I 最终汇总表
  # --------------------------------------------------------
  
  summary_table <- data.frame(
    
    Model = model_label,
    
    Formula = paste(
      deparse(
        formula_object
      ),
      collapse = " "
    ),
    
    N = n,
    
    Events = event_n,
    
    Event_Rate = event_n / n,
    
    Number_of_predictors =
      length(
        predictor_vars
      ),
    
    Number_of_imputations = m,
    
    Bootstrap_requested = B,
    
    Bootstrap_valid = valid_B,
    
    Bootstrap_failed =
      B -
      valid_B,
    
    Apparent_AUC =
      apparent_auc,
    
    Mean_Optimism_AUC =
      mean_optimism_auc,
    
    Optimism_Corrected_AUC =
      corrected_auc,
    
    Corrected_AUC_P2_5 =
      safe_quantile_13(
        bootstrap_detail$
          Corrected_AUC,
        0.025
      ),
    
    Corrected_AUC_P97_5 =
      safe_quantile_13(
        bootstrap_detail$
          Corrected_AUC,
        0.975
      ),
    
    Apparent_Brier =
      apparent_brier,
    
    Mean_Optimism_Brier =
      mean_optimism_brier,
    
    Optimism_Corrected_Brier =
      corrected_brier,
    
    Corrected_Brier_P2_5 =
      safe_quantile_13(
        bootstrap_detail$
          Corrected_Brier,
        0.025
      ),
    
    Corrected_Brier_P97_5 =
      safe_quantile_13(
        bootstrap_detail$
          Corrected_Brier,
        0.975
      ),
    
    Apparent_IPA =
      apparent_ipa,
    
    Optimism_Corrected_IPA =
      corrected_ipa,
    
    Apparent_LogLoss =
      apparent_logloss,
    
    Mean_Optimism_LogLoss =
      mean_optimism_logloss,
    
    Optimism_Corrected_LogLoss =
      corrected_logloss,
    
    Apparent_Calibration_Intercept =
      apparent_calibration_intercept,
    
    Mean_Optimism_Calibration_Intercept =
      mean_optimism_calibration_intercept,
    
    Optimism_Corrected_Calibration_Intercept =
      corrected_calibration_intercept,
    
    Corrected_Calibration_Intercept_P2_5 =
      safe_quantile_13(
        bootstrap_detail$
          Corrected_Calibration_Intercept,
        0.025
      ),
    
    Corrected_Calibration_Intercept_P97_5 =
      safe_quantile_13(
        bootstrap_detail$
          Corrected_Calibration_Intercept,
        0.975
      ),
    
    Apparent_Calibration_Slope =
      apparent_calibration_slope,
    
    Mean_Optimism_Calibration_Slope =
      mean_optimism_calibration_slope,
    
    Optimism_Corrected_Calibration_Slope =
      corrected_calibration_slope,
    
    Corrected_Calibration_Slope_P2_5 =
      safe_quantile_13(
        bootstrap_detail$
          Corrected_Calibration_Slope,
        0.025
      ),
    
    Corrected_Calibration_Slope_P97_5 =
      safe_quantile_13(
        bootstrap_detail$
          Corrected_Calibration_Slope,
        0.975
      ),
    
    Calibration_Intercept_Valid_B =
      sum(
        is.finite(
          bootstrap_detail$
            Optimism_Calibration_Intercept
        )
      ),
    
    Calibration_Slope_Valid_B =
      sum(
        is.finite(
          bootstrap_detail$
            Optimism_Calibration_Slope
        )
      ),
    
    Suggested_Global_Shrinkage_Factor =
      corrected_calibration_slope,
    
    Total_GLM_Warnings =
      total_fit_warnings,
    
    Validation_type =
      paste0(
        "MI-aware ordinary bootstrap internal validation; ",
        B,
        " resamples; final model formula fixed"
      ),
    
    stringsAsFactors = FALSE
  )
  
  # --------------------------------------------------------
  # 13.4J QC表
  # --------------------------------------------------------
  
  qc_table <- data.frame(
    
    Model = model_label,
    
    N = n,
    
    Events = event_n,
    
    Imputations = m,
    
    Bootstrap_requested = B,
    
    Bootstrap_valid = valid_B,
    
    Bootstrap_failed =
      B -
      valid_B,
    
    Valid_fraction =
      valid_B /
      B,
    
    Total_GLM_Warnings =
      total_fit_warnings,
    
    Apparent_AUC =
      apparent_auc,
    
    Corrected_AUC =
      corrected_auc,
    
    AUC_Optimism =
      mean_optimism_auc,
    
    Apparent_Brier =
      apparent_brier,
    
    Corrected_Brier =
      corrected_brier,
    
    Apparent_Calibration_Intercept =
      apparent_calibration_intercept,
    
    Corrected_Calibration_Intercept =
      corrected_calibration_intercept,
    
    Apparent_Calibration_Slope =
      apparent_calibration_slope,
    
    Corrected_Calibration_Slope =
      corrected_calibration_slope,
    
    Calibration_Intercept_Valid_B =
      sum(
        is.finite(
          bootstrap_detail$
            Optimism_Calibration_Intercept
        )
      ),
    
    Calibration_Slope_Valid_B =
      sum(
        is.finite(
          bootstrap_detail$
            Optimism_Calibration_Slope
        )
      ),
    
    stringsAsFactors = FALSE
  )
  
  cat(
    "\n",
    model_label,
    " Bootstrap验证完成：\n",
    "有效bootstrap = ",
    valid_B,
    "/",
    B,
    "\n",
    "AUC: apparent ",
    sprintf(
      "%.4f",
      apparent_auc
    ),
    " -> corrected ",
    sprintf(
      "%.4f",
      corrected_auc
    ),
    "\n",
    "Brier: apparent ",
    sprintf(
      "%.4f",
      apparent_brier
    ),
    " -> corrected ",
    sprintf(
      "%.4f",
      corrected_brier
    ),
    "\n",
    "Calibration intercept: ",
    sprintf(
      "%.4f",
      apparent_calibration_intercept
    ),
    " -> ",
    sprintf(
      "%.4f",
      corrected_calibration_intercept
    ),
    "\n",
    "Calibration slope: ",
    sprintf(
      "%.4f",
      apparent_calibration_slope
    ),
    " -> ",
    sprintf(
      "%.4f",
      corrected_calibration_slope
    ),
    "\n",
    sep = ""
  )
  
  list(
    model_label =
      model_label,
    formula =
      formula_object,
    apparent_by_mi =
      apparent_by_mi,
    apparent_prediction_mean =
      apparent_prediction_mean,
    bootstrap_detail =
      bootstrap_detail,
    failed_bootstrap_ids =
      failed_bootstrap_ids,
    summary =
      summary_table,
    qc =
      qc_table
  )
}


cat(
  "13.4 Bootstrap验证函数定义完成。\n"
)


# ==========================================================
# 13.5 运行术前模型Bootstrap
# ==========================================================

preoperative_bootstrap_13 <-
  bootstrap_validate_mi_model_13(
    
    imp_object =
      imp_preoperative,
    
    model_object =
      preoperative_final_model_12,
    
    model_label =
      "Preoperative",
    
    B =
      B_13,
    
    seed =
      seed_13,
    
    progress_every =
      progress_every_13
  )


# ==========================================================
# 13.6 运行拔管前模型Bootstrap
#
# 使用同一seed，从而在样本量一致时采用相同bootstrap患者索引。
# ==========================================================

precatheter_bootstrap_13 <-
  bootstrap_validate_mi_model_13(
    
    imp_object =
      imp_precatheter,
    
    model_object =
      precatheter_final_model_12,
    
    model_label =
      "Precatheter",
    
    B =
      B_13,
    
    seed =
      seed_13,
    
    progress_every =
      progress_every_13
  )


# ==========================================================
# 13.7 汇总两个最终模型
# ==========================================================

bootstrap_validation_summary_13 <-
  dplyr::bind_rows(
    
    preoperative_bootstrap_13$
      summary,
    
    precatheter_bootstrap_13$
      summary
  )


bootstrap_validation_qc_13 <-
  dplyr::bind_rows(
    
    preoperative_bootstrap_13$
      qc,
    
    precatheter_bootstrap_13$
      qc
  )


cat(
  "\n====================================================\n",
  "两个最终模型Bootstrap内部验证汇总：\n"
)

print(
  bootstrap_validation_summary_13
)


# ==========================================================
# 13.8 生成内部验证简洁性能汇总表
# ==========================================================

table3_internal_validation_13 <- data.frame(
  
  Model =
    bootstrap_validation_summary_13$
    Model,
  
  Apparent_AUC =
    round(
      bootstrap_validation_summary_13$
        Apparent_AUC,
      3
    ),
  
  Optimism_Corrected_AUC =
    round(
      bootstrap_validation_summary_13$
        Optimism_Corrected_AUC,
      3
    ),
  
  AUC_Optimism =
    round(
      bootstrap_validation_summary_13$
        Mean_Optimism_AUC,
      3
    ),
  
  Apparent_Brier =
    round(
      bootstrap_validation_summary_13$
        Apparent_Brier,
      4
    ),
  
  Optimism_Corrected_Brier =
    round(
      bootstrap_validation_summary_13$
        Optimism_Corrected_Brier,
      4
    ),
  
  Apparent_IPA =
    round(
      bootstrap_validation_summary_13$
        Apparent_IPA,
      4
    ),
  
  Optimism_Corrected_IPA =
    round(
      bootstrap_validation_summary_13$
        Optimism_Corrected_IPA,
      4
    ),
  
  Calibration_Intercept =
    round(
      bootstrap_validation_summary_13$
        Optimism_Corrected_Calibration_Intercept,
      3
    ),
  
  Calibration_Slope =
    round(
      bootstrap_validation_summary_13$
        Optimism_Corrected_Calibration_Slope,
      3
    ),
  
  Bootstrap_Valid =
    bootstrap_validation_summary_13$
    Bootstrap_valid,
  
  stringsAsFactors = FALSE
)


# ==========================================================
# 13.9 保存Excel结果
# ==========================================================

bootstrap_output_file_13 <- file.path(
  table_dir,
  "Bootstrap_internal_validation_summary.xlsx"
)


openxlsx::write.xlsx(
  
  list(
    
    Table3 =
      table3_internal_validation_13,
    
    Full_summary =
      bootstrap_validation_summary_13,
    
    Preoperative_bootstrap_detail =
      preoperative_bootstrap_13$
      bootstrap_detail,
    
    Precatheter_bootstrap_detail =
      precatheter_bootstrap_13$
      bootstrap_detail,
    
    Preoperative_apparent_by_MI =
      preoperative_bootstrap_13$
      apparent_by_mi,
    
    Precatheter_apparent_by_MI =
      precatheter_bootstrap_13$
      apparent_by_mi,
    
    QC =
      bootstrap_validation_qc_13
  ),
  
  bootstrap_output_file_13,
  
  overwrite = TRUE
)


# ==========================================================
# 13.10 保存QC日志
# ==========================================================

bootstrap_qc_file_13 <- file.path(
  log_dir,
  "Bootstrap_internal_validation_QC.xlsx"
)


openxlsx::write.xlsx(
  
  list(
    
    QC =
      bootstrap_validation_qc_13,
    
    Failed_bootstrap_Preoperative =
      data.frame(
        Bootstrap =
          preoperative_bootstrap_13$
          failed_bootstrap_ids
      ),
    
    Failed_bootstrap_Precatheter =
      data.frame(
        Bootstrap =
          precatheter_bootstrap_13$
          failed_bootstrap_ids
      )
  ),
  
  bootstrap_qc_file_13,
  
  overwrite = TRUE
)


# ==========================================================
# 13.11 保存Bootstrap对象
#
# 第14部分绘图时可直接读取：
# - apparent prediction
# - bootstrap detail
# - corrected performance
# ==========================================================

bootstrap_rds_file_13 <- file.path(
  model_dir,
  "Bootstrap_internal_validation_objects.rds"
)


saveRDS(
  
  list(
    
    Settings = list(
      B = B_13,
      seed = seed_13,
      validation =
        "MI-aware ordinary bootstrap; fixed final model formulas"
    ),
    
    Preoperative =
      preoperative_bootstrap_13,
    
    Precatheter =
      precatheter_bootstrap_13,
    
    Summary =
      bootstrap_validation_summary_13,
    
    Table3 =
      table3_internal_validation_13
  ),
  
  bootstrap_rds_file_13
)


# ==========================================================
# 13.12 最终质控
# ==========================================================

required_output_files_13 <- c(
  bootstrap_output_file_13,
  bootstrap_qc_file_13,
  bootstrap_rds_file_13
)


missing_output_files_13 <-
  required_output_files_13[
    !file.exists(
      required_output_files_13
    )
  ]


if (
  length(
    missing_output_files_13
  ) > 0
) {
  stop(
    paste0(
      "第13部分以下文件未成功生成：",
      paste(
        basename(
          missing_output_files_13
        ),
        collapse = ", "
      )
    )
  )
}


if (
  any(
    bootstrap_validation_qc_13$
    Valid_fraction < 0.95
  )
) {
  warning(
    "至少一个模型有效bootstrap比例低于95%，请重点检查QC文件。"
  )
}


if (
  any(
    !is.finite(
      bootstrap_validation_summary_13$
      Optimism_Corrected_AUC
    )
  )
) {
  stop(
    "至少一个模型无法得到有效的optimism-corrected AUC。"
  )
}


if (
  any(
    !is.finite(
      bootstrap_validation_summary_13$
      Optimism_Corrected_Brier
    )
  )
) {
  stop(
    "至少一个模型无法得到有效的optimism-corrected Brier。"
  )
}


cat(
  "\n>>> 第13部分全部完成：Bootstrap内部验证成功 <<<\n"
)


cat(
  "\n请重点检查：\n",
  "1. 02_tables/Bootstrap_internal_validation_summary.xlsx\n",
  "2. 05_logs/Bootstrap_internal_validation_QC.xlsx\n",
  "3. 03_models/Bootstrap_internal_validation_objects.rds\n",
  sep = ""
)


cat(
  "\n核心判读指标：\n",
  "- Optimism_Corrected_AUC：内部验证后的区分度\n",
  "- Optimism_Corrected_Brier：内部验证后的总体预测误差\n",
  "- Calibration_Intercept：理想值0\n",
  "- Calibration_Slope：理想值1；明显<1提示过拟合\n",
  "- Suggested_Global_Shrinkage_Factor：目前等于校正后的calibration slope，仅供后续裁决参考\n",
  sep = ""
)



############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
# 14. 最终模型可视化：ROC、Bootstrap bias-corrected Calibration、DCA
#
# 基于：
#   - 第12部分锁定的两个最终多因素Logistic模型
#   - 第13部分1000次Bootstrap内部验证结果
#
# 最终模型：
#   1. Preoperative model
#   2. Pre-catheter-removal model
#
# 本节输出：
#   Figure 1  ROC curves（双panel）
#   Figure 2  Calibration curves（双panel）
#   Figure 3  Decision Curve Analysis（双模型同图）
#   以及所有绘图数据、AUC CI和DCA数值表
#
# 重要方法学说明：
# 1. ROC曲线：
#    - 在20个MI数据集中分别生成ROC；
#    - 对共同FPR网格上的Sensitivity取平均，形成MI-averaged ROC曲线；
#    - AUC点估计为20个MI数据集AUC的均值，与第12/13部分口径一致；
#    - AUC 95%CI采用DeLong within-imputation variance +
#      Rubin-style between-imputation variance近似合并。
#    - 图内同时标记第13部分的optimism-corrected AUC。
#
# 2. Calibration曲线：
#    - 患者层面预测概率为20个MI模型预测概率的逐患者均值；
#    - 蓝点为按预测风险等人数分组后的 observed vs predicted；
#    - apparent curve 使用自然立方样条Logistic calibration smoother；
#    - 在1000次bootstrap中，对每次bootstrap样本与原始样本分别估计
#      flexible calibration curve，并逐点计算optimism；
#    - bias-corrected curve = apparent curve - mean bootstrap optimism；
#    - 因而本版solid line是真正的逐点bootstrap bias-corrected calibration curve，
#      不再用单一calibration intercept/slope代替整条校准曲线。
#
# 3. DCA：
#    - apparent DCA在20个MI数据集中分别计算Net Benefit后取均值；
#    - 重新执行与第13部分一致的1000次MI-aware bootstrap；
#    - 每个threshold分别计算：
#        optimism = NB_bootstrap sample - NB_original sample；
#      corrected NB = apparent NB - mean optimism；
#    - 正式Figure 3绘制optimism-corrected model curves；
#    - Treat all / Treat none不需要bootstrap校正；
#    - 主图阈值概率默认0.01-0.50，并保留Cost:Benefit Ratio刻度。
############################################################

cat("\n>>> 第14部分开始执行：ROC + Bootstrap-corrected Calibration + DCA <<<\n")
flush.console()


# ==========================================================
# 14.1 参数设置
# ==========================================================

# ROC平均曲线的FPR网格
roc_grid_n_14 <- 501L

# Calibration分组数
calibration_groups_14 <- 8L

# Calibration固定显示范围
calibration_xmax_14 <- 0.65

# Flexible calibration自然立方样条自由度
# 50个事件下df=3兼顾灵活性与稳定性
calibration_spline_df_14 <- 3L

# Calibration / DCA bootstrap设置
# 与第13部分保持一致；代码会进一步核对第13部分Bootstrap_requested。
curve_bootstrap_B_14 <- 1000L
curve_bootstrap_seed_14 <- 20260817L
curve_bootstrap_progress_every_14 <- 100L

# DCA阈值范围
DCA_threshold_min_14 <- 0.01
DCA_threshold_max_14 <- 0.50
DCA_threshold_step_14 <- 0.005

# 图像分辨率
figure_dpi_14 <- 600L

# 图中模型颜色（延续既往模板）
color_preoperative_14 <- "#0033CC"
color_precatheter_14 <- "#E41A1C"
color_all_14 <- "grey45"
color_none_14 <- "black"
color_point_14 <- "#0033CC"

# 概率裁剪
EPS_14 <- 1e-12


# ==========================================================
# 14.2 运行前检查 + 必要对象自动恢复
# ==========================================================

required_packages_14 <- c(
  "mice",
  "pROC",
  "dplyr",
  "openxlsx"
)

missing_packages_14 <- required_packages_14[
  !vapply(
    required_packages_14,
    requireNamespace,
    logical(1),
    quietly = TRUE
  )
]

if (length(missing_packages_14) > 0) {
  stop(
    paste0(
      "第14部分缺少R包：",
      paste(missing_packages_14, collapse = ", ")
    )
  )
}

required_base_objects_14 <- c(
  "imp_preoperative",
  "imp_precatheter",
  "outcome_var",
  "table_dir",
  "model_dir",
  "log_dir"
)

missing_base_objects_14 <- required_base_objects_14[
  !vapply(
    required_base_objects_14,
    exists,
    logical(1),
    inherits = TRUE
  )
]

if (length(missing_base_objects_14) > 0) {
  stop(
    paste0(
      "第14部分缺少前置对象：",
      paste(missing_base_objects_14, collapse = ", ")
    )
  )
}

# ----------------------------------------------------------
# 若第12部分最终模型对象不在当前Environment，尝试从RDS恢复
# ----------------------------------------------------------

if (
  !exists("preoperative_final_model_12", inherits = TRUE) ||
  !exists("precatheter_final_model_12", inherits = TRUE)
) {
  
  model_rds_14 <- file.path(
    model_dir,
    "Final_multivariable_model_objects.rds"
  )
  
  if (!file.exists(model_rds_14)) {
    stop(
      "找不到第12部分最终模型对象，也找不到 Final_multivariable_model_objects.rds。"
    )
  }
  
  model_objects_14 <- readRDS(model_rds_14)
  
  preoperative_final_model_12 <- model_objects_14$Preoperative_final
  precatheter_final_model_12 <- model_objects_14$Precatheter_final
}

# ----------------------------------------------------------
# 若第13部分Bootstrap对象不在当前Environment，尝试从RDS恢复
# ----------------------------------------------------------

if (
  !exists("preoperative_bootstrap_13", inherits = TRUE) ||
  !exists("precatheter_bootstrap_13", inherits = TRUE) ||
  !exists("bootstrap_validation_summary_13", inherits = TRUE)
) {
  
  bootstrap_rds_14 <- file.path(
    model_dir,
    "Bootstrap_internal_validation_objects.rds"
  )
  
  if (!file.exists(bootstrap_rds_14)) {
    stop(
      "找不到第13部分Bootstrap对象，也找不到 Bootstrap_internal_validation_objects.rds。"
    )
  }
  
  bootstrap_objects_14 <- readRDS(bootstrap_rds_14)
  
  preoperative_bootstrap_13 <- bootstrap_objects_14$Preoperative
  precatheter_bootstrap_13 <- bootstrap_objects_14$Precatheter
  bootstrap_validation_summary_13 <- bootstrap_objects_14$Summary
}

if (imp_preoperative$m != 20) {
  stop("术前MICE对象不是20个插补数据集。")
}

if (imp_precatheter$m != 20) {
  stop("拔管前MICE对象不是20个插补数据集。")
}

# ----------------------------------------------------------
# 图像目录：如前面已定义figure_dir则沿用，否则自动建立04_figures
# ----------------------------------------------------------

if (exists("figure_dir", inherits = TRUE)) {
  figure_dir_14 <- get("figure_dir", inherits = TRUE)
} else {
  figure_dir_14 <- file.path(
    dirname(table_dir),
    "04_figures"
  )
}

dir.create(
  figure_dir_14,
  recursive = TRUE,
  showWarnings = FALSE
)

cat(
  "14.2 运行前检查：通过。\n",
  "图像输出目录：",
  normalizePath(
    figure_dir_14,
    winslash = "/",
    mustWork = FALSE
  ),
  "\n",
  sep = ""
)


# ==========================================================
# 14.3 通用辅助函数
# ==========================================================

# ----------------------------------------------------------
# 14.3A 二分类结局安全转换
# ----------------------------------------------------------

coerce_binary_outcome_14 <- function(
    x,
    variable_name = "Outcome"
) {
  
  if (is.logical(x)) {
    x <- as.integer(x)
  }
  
  if (is.factor(x) || is.character(x)) {
    
    x_char <- trimws(as.character(x))
    
    valid_values <- unique(
      x_char[!is.na(x_char)]
    )
    
    if (!all(valid_values %in% c("0", "1"))) {
      stop(
        paste0(
          variable_name,
          " 不是明确的0/1编码。当前取值：",
          paste(valid_values, collapse = ", ")
        )
      )
    }
    
    x <- as.numeric(x_char)
  }
  
  x <- as.numeric(x)
  
  if (anyNA(x) || !all(x %in% c(0, 1))) {
    stop(
      paste0(
        variable_name,
        " 不是有效的0/1二分类结局。"
      )
    )
  }
  
  x
}


# ----------------------------------------------------------
# 14.3B 概率裁剪
# ----------------------------------------------------------

clip_probability_14 <- function(p) {
  pmin(
    pmax(as.numeric(p), EPS_14),
    1 - EPS_14
  )
}


# ----------------------------------------------------------
# 14.3C 从20个MI模型得到患者层面的预测概率矩阵
# ----------------------------------------------------------

get_mi_predictions_14 <- function(
    imp_object,
    model_object,
    model_label
) {
  
  m <- imp_object$m
  
  if (
    is.null(model_object$fits) ||
    length(model_object$fits) != m
  ) {
    stop(
      paste0(
        model_label,
        " 的最终模型对象中fits数量与MI数量不一致。"
      )
    )
  }
  
  formula_object <- model_object$formula
  
  if (is.null(formula_object)) {
    stop(
      paste0(
        model_label,
        " 缺少最终模型formula。"
      )
    )
  }
  
  predictor_vars <- attr(
    stats::terms(formula_object),
    "term.labels"
  )
  
  completed_list <- vector("list", m)
  pred_matrix <- NULL
  y_reference <- NULL
  
  for (i in seq_len(m)) {
    
    data_i <- mice::complete(
      imp_object,
      action = i
    )
    
    required_vars <- c(
      outcome_var,
      predictor_vars
    )
    
    missing_vars <- setdiff(
      required_vars,
      names(data_i)
    )
    
    if (length(missing_vars) > 0) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个MI数据集缺少变量：",
          paste(missing_vars, collapse = ", ")
        )
      )
    }
    
    if (anyNA(data_i[required_vars])) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个MI数据集最终模型变量仍有缺失。"
        )
      )
    }
    
    y_i <- coerce_binary_outcome_14(
      data_i[[outcome_var]],
      variable_name = outcome_var
    )
    
    if (is.null(y_reference)) {
      y_reference <- y_i
      pred_matrix <- matrix(
        NA_real_,
        nrow = nrow(data_i),
        ncol = m
      )
    } else {
      if (!identical(y_reference, y_i)) {
        stop(
          paste0(
            model_label,
            " 的MI数据集结局顺序不一致。"
          )
        )
      }
    }
    
    pred_i <- stats::predict(
      model_object$fits[[i]],
      newdata = data_i,
      type = "response"
    )
    
    if (
      length(pred_i) != length(y_reference) ||
      any(!is.finite(pred_i))
    ) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个模型预测概率无效。"
        )
      )
    }
    
    pred_matrix[, i] <- clip_probability_14(pred_i)
    completed_list[[i]] <- data_i
  }
  
  list(
    label = model_label,
    formula = formula_object,
    predictor_vars = predictor_vars,
    y = y_reference,
    prediction_matrix = pred_matrix,
    prediction_mean = rowMeans(pred_matrix),
    completed_list = completed_list
  )
}


# ----------------------------------------------------------
# 14.3D Rubin-style合并AUC + 95%CI
#
# 目的：图内AUC口径与第12/13部分“20个MI AUC均值”一致。
# Within-imputation variance采用DeLong。
# Between-imputation variance来自20个AUC之间的方差。
# ----------------------------------------------------------

pool_auc_mi_14 <- function(
    y,
    prediction_matrix,
    model_label
) {
  
  m <- ncol(prediction_matrix)
  
  auc_values <- rep(NA_real_, m)
  auc_variances <- rep(NA_real_, m)
  roc_objects <- vector("list", m)
  
  for (i in seq_len(m)) {
    
    roc_i <- pROC::roc(
      response = y,
      predictor = prediction_matrix[, i],
      levels = c(0, 1),
      direction = "<",
      quiet = TRUE
    )
    
    roc_objects[[i]] <- roc_i
    
    auc_values[i] <- as.numeric(
      pROC::auc(roc_i)
    )
    
    var_i <- tryCatch(
      as.numeric(
        pROC::var(
          roc_i,
          method = "delong"
        )
      ),
      error = function(e) NA_real_
    )
    
    auc_variances[i] <- var_i
  }
  
  if (any(!is.finite(auc_values))) {
    stop(
      paste0(
        model_label,
        " 存在无法计算的MI AUC。"
      )
    )
  }
  
  q_bar <- mean(auc_values)
  u_bar <- mean(auc_variances, na.rm = TRUE)
  b_var <- stats::var(auc_values)
  
  if (!is.finite(u_bar)) {
    u_bar <- 0
  }
  
  if (!is.finite(b_var)) {
    b_var <- 0
  }
  
  total_var <- u_bar +
    (1 + 1 / m) * b_var
  
  total_se <- sqrt(total_var)
  
  if (u_bar > 0 && b_var > 0) {
    
    r_value <- (
      (1 + 1 / m) * b_var
    ) / u_bar
    
    df_value <- (m - 1) *
      (1 + 1 / r_value)^2
    
    crit <- stats::qt(
      0.975,
      df = df_value
    )
    
  } else {
    
    df_value <- Inf
    crit <- stats::qnorm(0.975)
  }
  
  ci_lower <- max(
    0,
    q_bar - crit * total_se
  )
  
  ci_upper <- min(
    1,
    q_bar + crit * total_se
  )
  
  data.frame(
    Model = model_label,
    MI_n = m,
    Apparent_AUC = q_bar,
    AUC_SE = total_se,
    AUC_CI_lower = ci_lower,
    AUC_CI_upper = ci_upper,
    Rubin_df = df_value,
    Within_imputation_variance = u_bar,
    Between_imputation_variance = b_var,
    stringsAsFactors = FALSE
  ) -> summary_table
  
  list(
    summary = summary_table,
    auc_by_mi = data.frame(
      Model = model_label,
      Imputation = seq_len(m),
      AUC = auc_values,
      DeLong_variance = auc_variances,
      stringsAsFactors = FALSE
    ),
    roc_objects = roc_objects
  )
}


# ----------------------------------------------------------
# 14.3E MI-average ROC曲线
# ----------------------------------------------------------

average_roc_curve_14 <- function(
    roc_objects,
    model_label,
    grid_n = 501L
) {
  
  fpr_grid <- seq(
    0,
    1,
    length.out = grid_n
  )
  
  sensitivity_matrix <- matrix(
    NA_real_,
    nrow = grid_n,
    ncol = length(roc_objects)
  )
  
  for (i in seq_along(roc_objects)) {
    
    coords_i <- pROC::coords(
      roc_objects[[i]],
      x = "all",
      ret = c(
        "specificity",
        "sensitivity"
      ),
      transpose = FALSE
    )
    
    fpr_i <- 1 - as.numeric(
      coords_i$specificity
    )
    
    sens_i <- as.numeric(
      coords_i$sensitivity
    )
    
    # 按FPR升序，并处理重复FPR：取最大Sensitivity
    temp_i <- data.frame(
      FPR = fpr_i,
      Sensitivity = sens_i
    )
    
    temp_i <- temp_i[
      is.finite(temp_i$FPR) &
        is.finite(temp_i$Sensitivity),
      ,
      drop = FALSE
    ]
    
    temp_agg <- stats::aggregate(
      Sensitivity ~ FPR,
      data = temp_i,
      FUN = max
    )
    
    temp_agg <- temp_agg[
      order(temp_agg$FPR),
      ,
      drop = FALSE
    ]
    
    # 强制ROC端点
    if (temp_agg$FPR[1] > 0) {
      temp_agg <- rbind(
        data.frame(FPR = 0, Sensitivity = 0),
        temp_agg
      )
    }
    
    if (tail(temp_agg$FPR, 1) < 1) {
      temp_agg <- rbind(
        temp_agg,
        data.frame(FPR = 1, Sensitivity = 1)
      )
    }
    
    sensitivity_matrix[, i] <- stats::approx(
      x = temp_agg$FPR,
      y = temp_agg$Sensitivity,
      xout = fpr_grid,
      method = "linear",
      rule = 2,
      ties = "ordered"
    )$y
  }
  
  data.frame(
    Model = model_label,
    FPR = fpr_grid,
    Mean_Sensitivity = rowMeans(
      sensitivity_matrix,
      na.rm = TRUE
    ),
    stringsAsFactors = FALSE
  )
}


# ----------------------------------------------------------
# 14.3F Calibration分组点
# ----------------------------------------------------------

make_calibration_groups_14 <- function(
    y,
    p,
    model_label,
    groups = 8L
) {
  
  n <- length(y)
  
  if (groups < 3L) {
    stop("Calibration分组数至少为3。")
  }
  
  rank_index <- rank(
    p,
    ties.method = "first"
  )
  
  group_id <- ceiling(
    rank_index / n * groups
  )
  
  group_id[group_id < 1] <- 1
  group_id[group_id > groups] <- groups
  
  calibration_data <- data.frame(
    y = y,
    p = p,
    Group = group_id
  )
  
  calibration_data |>
    dplyr::group_by(Group) |>
    dplyr::summarise(
      Model = model_label,
      N = dplyr::n(),
      Events = sum(y == 1),
      Mean_predicted = mean(p),
      Observed_probability = mean(y),
      .groups = "drop"
    )
}


# ----------------------------------------------------------
# 14.3G Flexible calibration smoother
#
# 使用logit(predicted probability)上的自然立方样条Logistic模型：
#   logit(P(Y=1)) = f(logit(predicted risk))
#
# 该函数用于：
# - original sample apparent calibration curve
# - bootstrap-training calibration curve
# - bootstrap-test calibration curve
#
# 每条曲线都在同一个predicted-risk grid上评价，
# 因而可以逐点计算bootstrap optimism。
# ----------------------------------------------------------

fit_flexible_calibration_curve_14 <- function(
    y,
    p,
    x_grid,
    spline_df = 3L
) {
  
  y <- as.numeric(y)
  p <- clip_probability_14(p)
  x_grid <- clip_probability_14(x_grid)
  
  if (length(unique(y)) < 2) {
    return(
      list(
        valid = FALSE,
        curve = rep(NA_real_, length(x_grid)),
        warning_n = 0L,
        warning_text = "Outcome has only one level"
      )
    )
  }
  
  lp <- stats::qlogis(p)
  lp_grid <- stats::qlogis(x_grid)
  
  if (
    length(unique(round(lp, 10))) <
    max(5L, spline_df + 2L)
  ) {
    return(
      list(
        valid = FALSE,
        curve = rep(NA_real_, length(x_grid)),
        warning_n = 0L,
        warning_text = "Insufficient unique predicted risks"
      )
    )
  }
  
  captured_warnings <- character()
  
  fit_cal <- tryCatch(
    withCallingHandlers(
      stats::glm(
        y ~ splines::ns(lp, df = spline_df),
        family = stats::binomial()
      ),
      warning = function(w) {
        captured_warnings <<- c(
          captured_warnings,
          conditionMessage(w)
        )
        invokeRestart("muffleWarning")
      }
    ),
    error = function(e) NULL
  )
  
  if (
    is.null(fit_cal) ||
    !isTRUE(fit_cal$converged) ||
    any(!is.finite(stats::coef(fit_cal)))
  ) {
    return(
      list(
        valid = FALSE,
        curve = rep(NA_real_, length(x_grid)),
        warning_n = length(captured_warnings),
        warning_text = paste(unique(captured_warnings), collapse = " | ")
      )
    )
  }
  
  pred_curve <- tryCatch(
    stats::predict(
      fit_cal,
      newdata = data.frame(lp = lp_grid),
      type = "response"
    ),
    error = function(e) NULL
  )
  
  if (
    is.null(pred_curve) ||
    length(pred_curve) != length(x_grid) ||
    any(!is.finite(pred_curve))
  ) {
    return(
      list(
        valid = FALSE,
        curve = rep(NA_real_, length(x_grid)),
        warning_n = length(captured_warnings),
        warning_text = "Prediction from flexible calibration model failed"
      )
    )
  }
  
  list(
    valid = TRUE,
    curve = pmin(pmax(as.numeric(pred_curve), 0), 1),
    warning_n = length(captured_warnings),
    warning_text = if (
      length(captured_warnings) == 0
    ) {
      ""
    } else {
      paste(unique(captured_warnings), collapse = " | ")
    }
  )
}


# ----------------------------------------------------------
# 14.3H 单个概率向量的DCA Net Benefit
# ----------------------------------------------------------

net_benefit_vector_14 <- function(
    y,
    p,
    thresholds
) {
  
  y <- as.numeric(y)
  p <- clip_probability_14(p)
  n <- length(y)
  
  vapply(
    thresholds,
    function(pt) {
      
      predicted_positive <- p >= pt
      
      TP <- sum(
        predicted_positive & y == 1
      )
      
      FP <- sum(
        predicted_positive & y == 0
      )
      
      TP / n -
        FP / n *
        pt / (1 - pt)
    },
    numeric(1)
  )
}


# ----------------------------------------------------------
# 14.3I Apparent DCA：各MI分别计算后求均值
# ----------------------------------------------------------

calculate_apparent_dca_mi_14 <- function(
    y,
    prediction_matrix,
    model_label,
    thresholds
) {
  
  m <- ncol(prediction_matrix)
  
  nb_matrix <- matrix(
    NA_real_,
    nrow = m,
    ncol = length(thresholds)
  )
  
  for (i in seq_len(m)) {
    nb_matrix[i, ] <- net_benefit_vector_14(
      y = y,
      p = prediction_matrix[, i],
      thresholds = thresholds
    )
  }
  
  data.frame(
    Model = model_label,
    Threshold = thresholds,
    Apparent_Net_Benefit = colMeans(
      nb_matrix,
      na.rm = TRUE
    ),
    Apparent_NB_SD_across_MI = apply(
      nb_matrix,
      2,
      stats::sd,
      na.rm = TRUE
    ),
    stringsAsFactors = FALSE
  )
}


# ----------------------------------------------------------
# 14.3J Bootstrap逐点校正Calibration + DCA
#
# 每个bootstrap重复：
# 1. 使用同一患者bootstrap索引作用于20个MI数据集；
# 2. 在每个MI bootstrap sample重新拟合已锁定的最终模型公式；
# 3. 对bootstrap sample生成train predictions；
# 4. 对original sample生成test predictions；
# 5. Calibration：
#      - 20个MI预测逐患者平均；
#      - 分别拟合train/test flexible calibration curve；
#      - optimism_curve = train_curve - test_curve；
# 6. DCA：
#      - 每个MI分别计算train/test NB；
#      - 再跨MI求均值；
#      - optimism_NB = train_NB - test_NB；
# 7. 最终：
#      bias-corrected calibration = apparent - mean optimism_curve
#      corrected NB = apparent NB - mean optimism_NB
# ----------------------------------------------------------

bootstrap_correct_calibration_dca_14 <- function(
    imp_object,
    model_object,
    y,
    apparent_prediction_matrix,
    model_label,
    calibration_x_grid,
    thresholds,
    B = 1000L,
    seed = 20260817L,
    progress_every = 100L,
    spline_df = 3L
) {
  
  m <- imp_object$m
  n <- length(y)
  
  if (
    is.null(model_object$formula) ||
    is.null(model_object$fits) ||
    length(model_object$fits) != m
  ) {
    stop(
      paste0(
        model_label,
        " 的最终模型对象不完整。"
      )
    )
  }
  
  formula_object <- model_object$formula
  
  predictor_vars <- attr(
    stats::terms(formula_object),
    "term.labels"
  )
  
  completed_list <- vector("list", m)
  
  for (i in seq_len(m)) {
    
    data_i <- mice::complete(
      imp_object,
      action = i
    )
    
    required_vars <- c(
      outcome_var,
      predictor_vars
    )
    
    if (length(setdiff(required_vars, names(data_i))) > 0) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个MI数据集缺少最终模型变量。"
        )
      )
    }
    
    if (anyNA(data_i[required_vars])) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个MI数据集最终模型变量仍有缺失。"
        )
      )
    }
    
    y_i <- coerce_binary_outcome_14(
      data_i[[outcome_var]],
      variable_name = outcome_var
    )
    
    if (!identical(y, y_i)) {
      stop(
        paste0(
          model_label,
          " 第",
          i,
          "个MI数据集结局顺序与参考数据不一致。"
        )
      )
    }
    
    completed_list[[i]] <- data_i
  }
  
  apparent_prediction_mean <- rowMeans(
    apparent_prediction_matrix
  )
  
  apparent_calibration_fit <- fit_flexible_calibration_curve_14(
    y = y,
    p = apparent_prediction_mean,
    x_grid = calibration_x_grid,
    spline_df = spline_df
  )
  
  if (!isTRUE(apparent_calibration_fit$valid)) {
    stop(
      paste0(
        model_label,
        " 的apparent flexible calibration curve拟合失败。"
      )
    )
  }
  
  apparent_calibration_curve <-
    apparent_calibration_fit$curve
  
  apparent_dca <- calculate_apparent_dca_mi_14(
    y = y,
    prediction_matrix = apparent_prediction_matrix,
    model_label = model_label,
    thresholds = thresholds
  )
  
  set.seed(seed)
  
  bootstrap_indices <- replicate(
    B,
    sample.int(
      n = n,
      size = n,
      replace = TRUE
    ),
    simplify = FALSE
  )
  
  calibration_optimism_matrix <- matrix(
    NA_real_,
    nrow = B,
    ncol = length(calibration_x_grid)
  )
  
  dca_optimism_matrix <- matrix(
    NA_real_,
    nrow = B,
    ncol = length(thresholds)
  )
  
  model_fit_valid <- rep(FALSE, B)
  calibration_valid <- rep(FALSE, B)
  dca_valid <- rep(FALSE, B)
  bootstrap_events <- rep(NA_integer_, B)
  glm_warning_n <- integer(B)
  calibration_warning_n <- integer(B)
  
  failed_model_ids <- integer(0)
  failed_calibration_ids <- integer(0)
  
  cat(
    "\n开始Bootstrap curve correction：",
    model_label,
    "\n",
    sep = ""
  )
  
  for (b in seq_len(B)) {
    
    if (
      b == 1L ||
      b %% progress_every == 0L ||
      b == B
    ) {
      cat(
        "  Bootstrap ",
        b,
        "/",
        B,
        "\n",
        sep = ""
      )
      flush.console()
    }
    
    index_b <- bootstrap_indices[[b]]
    y_boot <- y[index_b]
    bootstrap_events[b] <- sum(y_boot == 1)
    
    if (length(unique(y_boot)) < 2) {
      failed_model_ids <- c(failed_model_ids, b)
      next
    }
    
    train_prediction_matrix <- matrix(
      NA_real_,
      nrow = n,
      ncol = m
    )
    
    test_prediction_matrix <- matrix(
      NA_real_,
      nrow = n,
      ncol = m
    )
    
    train_nb_mi <- matrix(
      NA_real_,
      nrow = m,
      ncol = length(thresholds)
    )
    
    test_nb_mi <- matrix(
      NA_real_,
      nrow = m,
      ncol = length(thresholds)
    )
    
    bootstrap_failed <- FALSE
    
    for (i in seq_len(m)) {
      
      data_i <- completed_list[[i]]
      
      data_boot <- data_i[
        index_b,
        ,
        drop = FALSE
      ]
      
      captured_warnings <- character()
      
      fit_boot <- tryCatch(
        withCallingHandlers(
          stats::glm(
            formula = formula_object,
            data = data_boot,
            family = stats::binomial()
          ),
          warning = function(w) {
            captured_warnings <<- c(
              captured_warnings,
              conditionMessage(w)
            )
            invokeRestart("muffleWarning")
          }
        ),
        error = function(e) NULL
      )
      
      glm_warning_n[b] <- glm_warning_n[b] +
        length(captured_warnings)
      
      if (
        is.null(fit_boot) ||
        !isTRUE(fit_boot$converged) ||
        any(!is.finite(stats::coef(fit_boot)))
      ) {
        bootstrap_failed <- TRUE
        break
      }
      
      pred_train_i <- tryCatch(
        stats::predict(
          fit_boot,
          newdata = data_boot,
          type = "response"
        ),
        error = function(e) NULL
      )
      
      pred_test_i <- tryCatch(
        stats::predict(
          fit_boot,
          newdata = data_i,
          type = "response"
        ),
        error = function(e) NULL
      )
      
      if (
        is.null(pred_train_i) ||
        is.null(pred_test_i) ||
        length(pred_train_i) != n ||
        length(pred_test_i) != n ||
        any(!is.finite(pred_train_i)) ||
        any(!is.finite(pred_test_i))
      ) {
        bootstrap_failed <- TRUE
        break
      }
      
      pred_train_i <- clip_probability_14(pred_train_i)
      pred_test_i <- clip_probability_14(pred_test_i)
      
      train_prediction_matrix[, i] <- pred_train_i
      test_prediction_matrix[, i] <- pred_test_i
      
      train_nb_mi[i, ] <- net_benefit_vector_14(
        y = y_boot,
        p = pred_train_i,
        thresholds = thresholds
      )
      
      test_nb_mi[i, ] <- net_benefit_vector_14(
        y = y,
        p = pred_test_i,
        thresholds = thresholds
      )
    }
    
    if (bootstrap_failed) {
      failed_model_ids <- c(failed_model_ids, b)
      next
    }
    
    model_fit_valid[b] <- TRUE
    
    # DCA逐threshold optimism
    train_nb_mean <- colMeans(
      train_nb_mi,
      na.rm = TRUE
    )
    
    test_nb_mean <- colMeans(
      test_nb_mi,
      na.rm = TRUE
    )
    
    dca_optimism_matrix[b, ] <-
      train_nb_mean - test_nb_mean
    
    dca_valid[b] <- all(
      is.finite(dca_optimism_matrix[b, ])
    )
    
    # Flexible calibration逐点optimism
    train_prediction_mean <- rowMeans(
      train_prediction_matrix
    )
    
    test_prediction_mean <- rowMeans(
      test_prediction_matrix
    )
    
    train_calibration_fit <- fit_flexible_calibration_curve_14(
      y = y_boot,
      p = train_prediction_mean,
      x_grid = calibration_x_grid,
      spline_df = spline_df
    )
    
    test_calibration_fit <- fit_flexible_calibration_curve_14(
      y = y,
      p = test_prediction_mean,
      x_grid = calibration_x_grid,
      spline_df = spline_df
    )
    
    calibration_warning_n[b] <-
      train_calibration_fit$warning_n +
      test_calibration_fit$warning_n
    
    if (
      isTRUE(train_calibration_fit$valid) &&
      isTRUE(test_calibration_fit$valid) &&
      all(is.finite(train_calibration_fit$curve)) &&
      all(is.finite(test_calibration_fit$curve))
    ) {
      
      calibration_optimism_matrix[b, ] <-
        train_calibration_fit$curve -
        test_calibration_fit$curve
      
      calibration_valid[b] <- TRUE
      
    } else {
      
      failed_calibration_ids <- c(
        failed_calibration_ids,
        b
      )
    }
  }
  
  # --------------------------------------------------------
  # Calibration correction
  # --------------------------------------------------------
  
  valid_calibration_B <- sum(calibration_valid)
  
  if (valid_calibration_B == 0L) {
    stop(
      paste0(
        model_label,
        " 没有任何有效的flexible calibration bootstrap重复。"
      )
    )
  }
  
  if (valid_calibration_B < 0.90 * B) {
    warning(
      paste0(
        model_label,
        " flexible calibration仅有 ",
        valid_calibration_B,
        "/",
        B,
        " 次有效bootstrap。"
      )
    )
  }
  
  mean_calibration_optimism <- apply(
    calibration_optimism_matrix[
      calibration_valid,
      ,
      drop = FALSE
    ],
    2,
    mean,
    na.rm = TRUE
  )
  
  bias_corrected_calibration_curve <-
    apparent_calibration_curve -
    mean_calibration_optimism
  
  bias_corrected_calibration_curve <- pmin(
    pmax(
      bias_corrected_calibration_curve,
      0
    ),
    1
  )
  
  corrected_calibration_matrix <- sweep(
    calibration_optimism_matrix[
      calibration_valid,
      ,
      drop = FALSE
    ],
    2,
    apparent_calibration_curve,
    FUN = function(opt, app) app - opt
  )
  
  corrected_calibration_matrix <- pmin(
    pmax(corrected_calibration_matrix, 0),
    1
  )
  
  calibration_curve_table <- data.frame(
    Model = model_label,
    Predicted = calibration_x_grid,
    Apparent = apparent_calibration_curve,
    Mean_Optimism = mean_calibration_optimism,
    Bias_corrected = bias_corrected_calibration_curve,
    Bias_corrected_P2_5 = apply(
      corrected_calibration_matrix,
      2,
      stats::quantile,
      probs = 0.025,
      na.rm = TRUE,
      names = FALSE
    ),
    Bias_corrected_P97_5 = apply(
      corrected_calibration_matrix,
      2,
      stats::quantile,
      probs = 0.975,
      na.rm = TRUE,
      names = FALSE
    ),
    Ideal = calibration_x_grid,
    stringsAsFactors = FALSE
  )
  
  # --------------------------------------------------------
  # DCA correction
  # --------------------------------------------------------
  
  valid_dca_B <- sum(dca_valid)
  
  if (valid_dca_B == 0L) {
    stop(
      paste0(
        model_label,
        " 没有任何有效的DCA bootstrap重复。"
      )
    )
  }
  
  if (valid_dca_B < 0.95 * B) {
    warning(
      paste0(
        model_label,
        " DCA仅有 ",
        valid_dca_B,
        "/",
        B,
        " 次有效bootstrap。"
      )
    )
  }
  
  mean_dca_optimism <- apply(
    dca_optimism_matrix[
      dca_valid,
      ,
      drop = FALSE
    ],
    2,
    mean,
    na.rm = TRUE
  )
  
  corrected_dca <-
    apparent_dca$Apparent_Net_Benefit -
    mean_dca_optimism
  
  corrected_dca_matrix <- sweep(
    dca_optimism_matrix[
      dca_valid,
      ,
      drop = FALSE
    ],
    2,
    apparent_dca$Apparent_Net_Benefit,
    FUN = function(opt, app) app - opt
  )
  
  dca_table <- data.frame(
    Model = model_label,
    Threshold = thresholds,
    Apparent_Net_Benefit =
      apparent_dca$Apparent_Net_Benefit,
    Apparent_NB_SD_across_MI =
      apparent_dca$Apparent_NB_SD_across_MI,
    Mean_Optimism_Net_Benefit =
      mean_dca_optimism,
    Optimism_Corrected_Net_Benefit =
      corrected_dca,
    Corrected_NB_P2_5 = apply(
      corrected_dca_matrix,
      2,
      stats::quantile,
      probs = 0.025,
      na.rm = TRUE,
      names = FALSE
    ),
    Corrected_NB_P97_5 = apply(
      corrected_dca_matrix,
      2,
      stats::quantile,
      probs = 0.975,
      na.rm = TRUE,
      names = FALSE
    ),
    Bootstrap_valid = valid_dca_B,
    stringsAsFactors = FALSE
  )
  
  qc_table <- data.frame(
    Model = model_label,
    Bootstrap_requested = B,
    Model_fit_valid_B = sum(model_fit_valid),
    Model_fit_failed_B = B - sum(model_fit_valid),
    Calibration_valid_B = valid_calibration_B,
    DCA_valid_B = valid_dca_B,
    Total_GLM_warnings = sum(glm_warning_n),
    Total_calibration_warnings = sum(calibration_warning_n),
    Spline_df = spline_df,
    Calibration_x_min = min(calibration_x_grid),
    Calibration_x_max = max(calibration_x_grid),
    DCA_threshold_min = min(thresholds),
    DCA_threshold_max = max(thresholds),
    stringsAsFactors = FALSE
  )
  
  list(
    calibration_curve = calibration_curve_table,
    dca = dca_table,
    qc = qc_table,
    failed_model_ids = failed_model_ids,
    failed_calibration_ids = failed_calibration_ids
  )
}


# 14.4 生成两个模型的MI预测
# ==========================================================

preoperative_plot_14 <- get_mi_predictions_14(
  imp_object = imp_preoperative,
  model_object = preoperative_final_model_12,
  model_label = "Preoperative model"
)

precatheter_plot_14 <- get_mi_predictions_14(
  imp_object = imp_precatheter,
  model_object = precatheter_final_model_12,
  model_label = "Pre-catheter-removal model"
)

if (!identical(
  preoperative_plot_14$y,
  precatheter_plot_14$y
)) {
  stop(
    "术前模型与拔管前模型的结局顺序不一致，无法直接比较ROC/DCA。"
  )
}

if (
  length(preoperative_plot_14$y) !=
  length(precatheter_plot_14$y)
) {
  stop(
    "术前模型与拔管前模型样本量不一致。"
  )
}

y_14 <- preoperative_plot_14$y
N_14 <- length(y_14)
Events_14 <- sum(y_14 == 1)
Prevalence_14 <- mean(y_14)

cat(
  "14.4 MI预测生成完成：N = ",
  N_14,
  ", Events = ",
  Events_14,
  ", POUR rate = ",
  sprintf("%.2f%%", 100 * Prevalence_14),
  "\n",
  sep = ""
)


# ==========================================================
# 14.5 ROC数据 + AUC 95%CI
# ==========================================================

preoperative_auc_14 <- pool_auc_mi_14(
  y = y_14,
  prediction_matrix = preoperative_plot_14$prediction_matrix,
  model_label = "Preoperative model"
)

precatheter_auc_14 <- pool_auc_mi_14(
  y = y_14,
  prediction_matrix = precatheter_plot_14$prediction_matrix,
  model_label = "Pre-catheter-removal model"
)

preoperative_roc_curve_14 <- average_roc_curve_14(
  roc_objects = preoperative_auc_14$roc_objects,
  model_label = "Preoperative model",
  grid_n = roc_grid_n_14
)

precatheter_roc_curve_14 <- average_roc_curve_14(
  roc_objects = precatheter_auc_14$roc_objects,
  model_label = "Pre-catheter-removal model",
  grid_n = roc_grid_n_14
)

roc_auc_summary_14 <- dplyr::bind_rows(
  preoperative_auc_14$summary,
  precatheter_auc_14$summary
)

roc_auc_by_mi_14 <- dplyr::bind_rows(
  preoperative_auc_14$auc_by_mi,
  precatheter_auc_14$auc_by_mi
)

# ----------------------------------------------------------
# 提取第13部分optimism-corrected AUC
# ----------------------------------------------------------

summary_preop_13_14 <- bootstrap_validation_summary_13[
  bootstrap_validation_summary_13$Model == "Preoperative",
  ,
  drop = FALSE
]

summary_precat_13_14 <- bootstrap_validation_summary_13[
  bootstrap_validation_summary_13$Model == "Precatheter",
  ,
  drop = FALSE
]

if (
  nrow(summary_preop_13_14) != 1 ||
  nrow(summary_precat_13_14) != 1
) {
  stop(
    "第13部分Bootstrap summary中无法唯一找到Preoperative / Precatheter结果。"
  )
}

roc_auc_summary_14$Optimism_Corrected_AUC <- c(
  summary_preop_13_14$Optimism_Corrected_AUC,
  summary_precat_13_14$Optimism_Corrected_AUC
)

cat("\nROC/AUC汇总：\n")
print(roc_auc_summary_14)


# ==========================================================
# 14.6 Calibration + DCA：1000次Bootstrap逐点optimism校正
#
# 说明：
# 第13部分保存了AUC/Brier/calibration intercept/slope的bootstrap结果，
# 但没有保存每个bootstrap的患者级train/test predictions，
# 因而无法仅凭第13部分对象重建“整条”bias-corrected calibration curve
# 或逐threshold的optimism-corrected DCA。
#
# 本节因此使用与第13部分相同的：
# - 最终模型公式
# - 20个MI数据集
# - bootstrap次数
# - 随机种子
# 重新执行bootstrap，仅用于曲线级校正。
# ==========================================================

if (
  length(unique(c(
    summary_preop_13_14$Bootstrap_requested,
    summary_precat_13_14$Bootstrap_requested
  ))) != 1
) {
  stop("第13部分两个模型的Bootstrap_requested不一致。")
}

curve_bootstrap_B_14 <- as.integer(
  summary_preop_13_14$Bootstrap_requested
)

if (curve_bootstrap_B_14 != 1000L) {
  warning(
    paste0(
      "第13部分Bootstrap次数为 ",
      curve_bootstrap_B_14,
      "，第14部分将按该次数进行曲线校正。"
    )
  )
}

# Calibration grid覆盖当前实际最大预测风险；
# 0.65可完整覆盖本数据约0.62的最大预测概率。
calibration_x_grid_14 <- seq(
  0.001,
  calibration_xmax_14,
  length.out = 300L
)

thresholds_14 <- seq(
  DCA_threshold_min_14,
  DCA_threshold_max_14,
  by = DCA_threshold_step_14
)

preoperative_curve_validation_14 <-
  bootstrap_correct_calibration_dca_14(
    imp_object = imp_preoperative,
    model_object = preoperative_final_model_12,
    y = y_14,
    apparent_prediction_matrix =
      preoperative_plot_14$prediction_matrix,
    model_label = "Preoperative model",
    calibration_x_grid = calibration_x_grid_14,
    thresholds = thresholds_14,
    B = curve_bootstrap_B_14,
    seed = curve_bootstrap_seed_14,
    progress_every = curve_bootstrap_progress_every_14,
    spline_df = calibration_spline_df_14
  )

precatheter_curve_validation_14 <-
  bootstrap_correct_calibration_dca_14(
    imp_object = imp_precatheter,
    model_object = precatheter_final_model_12,
    y = y_14,
    apparent_prediction_matrix =
      precatheter_plot_14$prediction_matrix,
    model_label = "Pre-catheter-removal model",
    calibration_x_grid = calibration_x_grid_14,
    thresholds = thresholds_14,
    B = curve_bootstrap_B_14,
    seed = curve_bootstrap_seed_14,
    progress_every = curve_bootstrap_progress_every_14,
    spline_df = calibration_spline_df_14
  )


# ==========================================================
# 14.7 整理Calibration与DCA正式绘图数据
# ==========================================================

preoperative_calibration_groups_14 <-
  make_calibration_groups_14(
    y = y_14,
    p = preoperative_plot_14$prediction_mean,
    model_label = "Preoperative model",
    groups = calibration_groups_14
  )

precatheter_calibration_groups_14 <-
  make_calibration_groups_14(
    y = y_14,
    p = precatheter_plot_14$prediction_mean,
    model_label = "Pre-catheter-removal model",
    groups = calibration_groups_14
  )

preoperative_calibration_14 <- list(
  curve = preoperative_curve_validation_14$calibration_curve,
  groups = preoperative_calibration_groups_14,
  metrics = data.frame(
    Model = "Preoperative model",
    N = N_14,
    Events = Events_14,
    Bootstrap_B = curve_bootstrap_B_14,
    Corrected_intercept =
      summary_preop_13_14$Optimism_Corrected_Calibration_Intercept,
    Corrected_slope =
      summary_preop_13_14$Optimism_Corrected_Calibration_Slope,
    stringsAsFactors = FALSE
  )
)

precatheter_calibration_14 <- list(
  curve = precatheter_curve_validation_14$calibration_curve,
  groups = precatheter_calibration_groups_14,
  metrics = data.frame(
    Model = "Pre-catheter-removal model",
    N = N_14,
    Events = Events_14,
    Bootstrap_B = curve_bootstrap_B_14,
    Corrected_intercept =
      summary_precat_13_14$Optimism_Corrected_Calibration_Intercept,
    Corrected_slope =
      summary_precat_13_14$Optimism_Corrected_Calibration_Slope,
    stringsAsFactors = FALSE
  )
)

calibration_metrics_14 <- dplyr::bind_rows(
  preoperative_calibration_14$metrics,
  precatheter_calibration_14$metrics
)

preoperative_dca_14 <-
  preoperative_curve_validation_14$dca

precatheter_dca_14 <-
  precatheter_curve_validation_14$dca

# Treat all / none不涉及模型估计，因此无需optimism correction
DCA_treat_all_14 <- data.frame(
  Model = "Treat all",
  Threshold = thresholds_14,
  Apparent_Net_Benefit = Prevalence_14 -
    (1 - Prevalence_14) *
    thresholds_14 /
    (1 - thresholds_14),
  Mean_Optimism_Net_Benefit = 0,
  Optimism_Corrected_Net_Benefit = Prevalence_14 -
    (1 - Prevalence_14) *
    thresholds_14 /
    (1 - thresholds_14),
  stringsAsFactors = FALSE
)

DCA_treat_none_14 <- data.frame(
  Model = "Treat none",
  Threshold = thresholds_14,
  Apparent_Net_Benefit = 0,
  Mean_Optimism_Net_Benefit = 0,
  Optimism_Corrected_Net_Benefit = 0,
  stringsAsFactors = FALSE
)

DCA_all_14 <- dplyr::bind_rows(
  preoperative_dca_14 |>
    dplyr::select(
      Model,
      Threshold,
      Apparent_Net_Benefit,
      Mean_Optimism_Net_Benefit,
      Optimism_Corrected_Net_Benefit
    ),
  precatheter_dca_14 |>
    dplyr::select(
      Model,
      Threshold,
      Apparent_Net_Benefit,
      Mean_Optimism_Net_Benefit,
      Optimism_Corrected_Net_Benefit
    ),
  DCA_treat_all_14,
  DCA_treat_none_14
)

curve_bootstrap_qc_14 <- dplyr::bind_rows(
  preoperative_curve_validation_14$qc,
  precatheter_curve_validation_14$qc
)

cat("\nCalibration curve / DCA bootstrap QC：\n")
print(curve_bootstrap_qc_14)

# ==========================================================
# 14.8 绘图辅助函数（最终投稿版V3）
# ==========================================================

# 14.8A ROC双panel绘图
# ----------------------------------------------------------

plot_roc_panels_14 <- function() {
  
  old_par <- par(no.readonly = TRUE)
  on.exit(par(old_par), add = TRUE)
  
  par(
    mfrow = c(1, 2),
    mar = c(4.9, 5.0, 3.4, 1.2),
    oma = c(0.2, 0.2, 0.2, 0.2),
    family = "sans",
    cex.axis = 0.92,
    cex.lab = 1.05,
    las = 1
  )
  
  draw_one_roc_14 <- function(
    roc_curve,
    auc_summary,
    corrected_auc,
    line_color,
    panel_label,
    panel_title
  ) {
    
    plot(
      roc_curve$FPR,
      roc_curve$Mean_Sensitivity,
      type = "n",
      xlim = c(0, 1),
      ylim = c(0, 1),
      xlab = "1 - Specificity",
      ylab = "Sensitivity",
      main = "",
      xaxs = "i",
      yaxs = "i",
      axes = FALSE
    )
    
    ticks_roc <- seq(0, 1, 0.25)
    
    axis(
      1,
      at = ticks_roc,
      labels = sprintf("%.2f", ticks_roc),
      tck = -0.015
    )
    
    axis(
      2,
      at = ticks_roc,
      labels = sprintf("%.2f", ticks_roc),
      tck = -0.015
    )
    
    box(lwd = 1.0)
    
    abline(
      h = ticks_roc[2:4],
      v = ticks_roc[2:4],
      col = "grey92",
      lwd = 0.8
    )
    
    abline(
      a = 0,
      b = 1,
      col = "grey75",
      lwd = 1.0
    )
    
    lines(
      roc_curve$FPR,
      roc_curve$Mean_Sensitivity,
      lwd = 2.2,
      col = line_color
    )
    
    title(
      main = panel_title,
      line = 1.2,
      cex.main = 1.05,
      font.main = 2
    )
    
    mtext(
      panel_label,
      side = 3,
      line = 1.15,
      adj = 0,
      font = 2,
      cex = 1.25
    )
    
    line1 <- paste0(
      "Apparent AUC = ",
      sprintf("%.3f", auc_summary$Apparent_AUC),
      " (95% CI, ",
      sprintf("%.3f", auc_summary$AUC_CI_lower),
      "-",
      sprintf("%.3f", auc_summary$AUC_CI_upper),
      ")"
    )
    
    line2 <- paste0(
      "Optimism-corrected AUC = ",
      sprintf("%.3f", corrected_auc)
    )
    
    legend(
      "bottomright",
      legend = c(line1, line2),
      bty = "o",
      bg = "white",
      box.col = "grey70",
      text.col = "black",
      cex = 0.74,
      x.intersp = 0.4,
      y.intersp = 1.15,
      inset = 0.02
    )
  }
  
  draw_one_roc_14(
    roc_curve = preoperative_roc_curve_14,
    auc_summary = preoperative_auc_14$summary,
    corrected_auc = summary_preop_13_14$Optimism_Corrected_AUC,
    line_color = color_preoperative_14,
    panel_label = "A",
    panel_title = "Preoperative model"
  )
  
  draw_one_roc_14(
    roc_curve = precatheter_roc_curve_14,
    auc_summary = precatheter_auc_14$summary,
    corrected_auc = summary_precat_13_14$Optimism_Corrected_AUC,
    line_color = color_precatheter_14,
    panel_label = "B",
    panel_title = "Pre-catheter-removal model"
  )
}


# ----------------------------------------------------------
# 14.8B Calibration双panel绘图
#
# 现在：
# dotted = apparent flexible calibration curve
# solid  = pointwise bootstrap bias-corrected calibration curve
# dashed = ideal
# ----------------------------------------------------------

plot_calibration_panels_14 <- function() {
  
  old_par <- par(no.readonly = TRUE)
  on.exit(par(old_par), add = TRUE)
  
  par(
    mfrow = c(1, 2),
    mar = c(6.2, 5.0, 3.6, 1.2),
    oma = c(0.2, 0.2, 0.2, 0.2),
    family = "sans",
    cex.axis = 0.92,
    cex.lab = 1.05,
    las = 1
  )
  
  draw_one_calibration_14 <- function(
    calibration_object,
    prediction_values,
    panel_label,
    panel_title
  ) {
    
    curve_data <- calibration_object$curve
    group_data <- calibration_object$groups
    metrics <- calibration_object$metrics
    x_max <- calibration_xmax_14
    
    plot(
      NA,
      xlim = c(0, x_max),
      ylim = c(0, x_max),
      xlab = "Predicted Probability",
      ylab = "Observed Probability",
      main = "",
      xaxs = "i",
      yaxs = "i",
      axes = FALSE
    )
    
    ticks <- seq(0, 0.60, 0.10)
    
    axis(
      1,
      at = ticks,
      labels = sprintf("%.1f", ticks),
      tck = -0.015
    )
    
    axis(
      2,
      at = ticks,
      labels = sprintf("%.1f", ticks),
      tck = -0.015
    )
    
    box(lwd = 1.0)
    
    inner_ticks <- ticks[-1]
    
    abline(
      h = inner_ticks,
      v = inner_ticks,
      col = "grey92",
      lwd = 0.8
    )
    
    # Ideal
    abline(
      a = 0,
      b = 1,
      lty = 2,
      lwd = 1.6,
      col = "black"
    )
    
    # Apparent flexible curve
    lines(
      curve_data$Predicted,
      curve_data$Apparent,
      lty = 3,
      lwd = 1.7,
      col = "black"
    )
    
    # 真正逐点bootstrap bias-corrected flexible curve
    lines(
      curve_data$Predicted,
      curve_data$Bias_corrected,
      lty = 1,
      lwd = 2.0,
      col = "black"
    )
    
    # 分组观察点仅辅助展示，不参与bootstrap curve correction本身
    points(
      group_data$Mean_predicted,
      group_data$Observed_probability,
      pch = 19,
      cex = 0.72,
      col = color_point_14
    )
    
    # 完整显示0-0.65范围内所有患者预测风险
    rug(
      prediction_values[
        prediction_values <= x_max
      ],
      side = 3,
      ticksize = 0.025,
      col = "grey35"
    )
    
    title(
      main = panel_title,
      line = 1.25,
      cex.main = 1.00,
      font.main = 2
    )
    
    mtext(
      panel_label,
      side = 3,
      line = 1.15,
      adj = 0,
      font = 2,
      cex = 1.25
    )
    
    legend(
      "bottomright",
      legend = c(
        "Apparent",
        "Bias-corrected",
        "Ideal"
      ),
      lty = c(3, 1, 2),
      lwd = c(1.7, 2.0, 1.6),
      col = c("black", "black", "black"),
      bty = "n",
      cex = 0.76,
      seg.len = 2.2,
      y.intersp = 1.0,
      inset = 0.02
    )
    
    mtext(
      paste0(
        "B = ",
        metrics$Bootstrap_B,
        " bootstrap repetitions"
      ),
      side = 1,
      line = 4.2,
      adj = 0,
      cex = 0.62
    )
    
    mtext(
      paste0(
        "n = ",
        metrics$N
      ),
      side = 1,
      line = 4.2,
      adj = 1,
      cex = 0.62
    )
  }
  
  draw_one_calibration_14(
    calibration_object = preoperative_calibration_14,
    prediction_values = preoperative_plot_14$prediction_mean,
    panel_label = "A",
    panel_title = "Calibration curve - Preoperative model"
  )
  
  draw_one_calibration_14(
    calibration_object = precatheter_calibration_14,
    prediction_values = precatheter_plot_14$prediction_mean,
    panel_label = "B",
    panel_title = "Calibration curve - Pre-catheter-removal model"
  )
}


# ----------------------------------------------------------
# 14.8C DCA绘图
#
# 正文主图聚焦临床相关threshold：
# X轴：0-0.25
# Y轴：-0.01-0.085
#
# 注意：
# 数据仍然来自完整0.01-0.50范围的
# 1000次bootstrap optimism-corrected DCA，
#这里只改变显示范围，不改变任何计算结果。
# ----------------------------------------------------------

plot_dca_14 <- function() {
  
  old_par <- par(no.readonly = TRUE)
  on.exit(par(old_par), add = TRUE)
  
  
  # ========================================================
  # 正文DCA显示范围
  # ========================================================
  
  dca_display_xmin_14 <- 0.00
  dca_display_xmax_14 <- 0.25
  
  dca_display_ymin_14 <- -0.01
  dca_display_ymax_14 <- 0.085
  
  
  par(
    mar = c(7.4, 5.4, 3.8, 1.2),
    family = "sans",
    cex.axis = 0.92,
    cex.lab = 1.05,
    las = 1
  )
  
  
  # ========================================================
  # 建立空白画布
  # ========================================================
  
  plot(
    preoperative_dca_14$Threshold,
    preoperative_dca_14$
      Optimism_Corrected_Net_Benefit,
    
    type = "n",
    
    xlim = c(
      dca_display_xmin_14,
      dca_display_xmax_14
    ),
    
    ylim = c(
      dca_display_ymin_14,
      dca_display_ymax_14
    ),
    
    xlab = "",
    ylab = "Net Benefit",
    
    main = "Decision Curve Analysis",
    
    xaxs = "i",
    yaxs = "i",
    
    axes = FALSE,
    
    cex.main = 1.18
  )
  
  
  # ========================================================
  # X轴：Threshold probability
  # ========================================================
  
  threshold_ticks_14 <- seq(
    0,
    0.25,
    by = 0.05
  )
  
  
  axis(
    side = 1,
    
    at = threshold_ticks_14,
    
    labels = sprintf(
      "%.2f",
      threshold_ticks_14
    ),
    
    tck = -0.015,
    
    mgp = c(
      3,
      0.7,
      0
    )
  )
  
  
  # ========================================================
  # Y轴：Net benefit
  # ========================================================
  
  net_benefit_ticks_14 <- c(
    -0.01,
    0.00,
    0.02,
    0.04,
    0.06,
    0.08
  )
  
  
  axis(
    side = 2,
    
    at = net_benefit_ticks_14,
    
    labels = sprintf(
      "%.2f",
      net_benefit_ticks_14
    ),
    
    tck = -0.015
  )
  
  
  # ========================================================
  # 边框
  # ========================================================
  
  box(
    lwd = 1.0
  )
  
  
  # ========================================================
  # 浅色辅助网格
  # ========================================================
  
  abline(
    h = c(
      0.02,
      0.04,
      0.06,
      0.08
    ),
    
    v = c(
      0.05,
      0.10,
      0.15,
      0.20
    ),
    
    col = "grey90",
    
    lwd = 0.8
  )
  
  
  # ========================================================
  # Treat all
  #
  # 超出Y轴下界的部分自然被裁剪。
  # 这是为了避免Treat-all在较高threshold的巨大负值
  # 压缩真正有临床意义的模型曲线。
  # ========================================================
  
  lines(
    DCA_treat_all_14$Threshold,
    DCA_treat_all_14$
      Optimism_Corrected_Net_Benefit,
    
    lwd = 1.5,
    
    col = color_all_14
  )
  
  
  # ========================================================
  # Treat none
  # ========================================================
  
  abline(
    h = 0,
    
    lwd = 1.5,
    
    col = color_none_14
  )
  
  
  # ========================================================
  # Preoperative model
  # ========================================================
  
  lines(
    preoperative_dca_14$Threshold,
    preoperative_dca_14$
      Optimism_Corrected_Net_Benefit,
    
    lwd = 3.0,
    
    col = color_preoperative_14
  )
  
  
  # ========================================================
  # Pre-catheter-removal model
  # ========================================================
  
  lines(
    precatheter_dca_14$Threshold,
    precatheter_dca_14$
      Optimism_Corrected_Net_Benefit,
    
    lwd = 3.0,
    
    col = color_precatheter_14
  )
  
  
  # ========================================================
  # 图例
  # ========================================================
  
  legend(
    "topright",
    
    legend = c(
      "Preoperative model",
      "Pre-catheter-removal model",
      "Treat all",
      "Treat none"
    ),
    
    col = c(
      color_preoperative_14,
      color_precatheter_14,
      color_all_14,
      color_none_14
    ),
    
    lwd = c(
      3.0,
      3.0,
      1.5,
      1.5
    ),
    
    bty = "o",
    
    bg = "white",
    
    box.col = "grey35",
    
    cex = 0.78,
    
    y.intersp = 1.05,
    
    inset = 0.018
  )
  
  
  # ========================================================
  # 第一层横轴标题
  # ========================================================
  
  mtext(
    "Threshold Probability",
    
    side = 1,
    
    line = 2.35,
    
    cex = 1.02
  )
  
  
  # ========================================================
  # 第二层：Cost:Benefit ratio
  #
  # odds = pt / (1 - pt)
  # ========================================================
  
  cb_at_14 <- c(
    0.01,
    0.05,
    0.10,
    0.15,
    0.20,
    0.25
  )
  
  
  cb_labels_14 <- c(
    "1:99",
    "1:19",
    "1:9",
    "3:17",
    "1:4",
    "1:3"
  )
  
  
  axis(
    side = 1,
    
    at = cb_at_14,
    
    labels = cb_labels_14,
    
    line = 3.65,
    
    tick = TRUE,
    
    cex.axis = 0.77
  )
  
  
  mtext(
    "Cost:Benefit Ratio",
    
    side = 1,
    
    line = 5.05,
    
    cex = 0.95
  )
}


# 14.9 保存Figure 1：ROC
# ==========================================================

roc_png_14 <- file.path(
  figure_dir_14,
  "Figure_ROC_final_models.png"
)

roc_tiff_14 <- file.path(
  figure_dir_14,
  "Figure_ROC_final_models.tiff"
)

roc_pdf_14 <- file.path(
  figure_dir_14,
  "Figure_ROC_final_models.pdf"
)

png(
  filename = roc_png_14,
  width = 12,
  height = 6.3,
  units = "in",
  res = figure_dpi_14,
  bg = "white"
)
plot_roc_panels_14()
dev.off()

tiff(
  filename = roc_tiff_14,
  width = 12,
  height = 6.3,
  units = "in",
  res = figure_dpi_14,
  compression = "lzw",
  bg = "white"
)
plot_roc_panels_14()
dev.off()

pdf(
  file = roc_pdf_14,
  width = 12,
  height = 6.3,
  family = "sans"
)
plot_roc_panels_14()
dev.off()

cat("14.9 ROC图保存完成。\n")


# ==========================================================
# 14.10 保存Figure 2：Calibration
# ==========================================================

cal_png_14 <- file.path(
  figure_dir_14,
  "Figure_Calibration_final_models.png"
)

cal_tiff_14 <- file.path(
  figure_dir_14,
  "Figure_Calibration_final_models.tiff"
)

cal_pdf_14 <- file.path(
  figure_dir_14,
  "Figure_Calibration_final_models.pdf"
)

png(
  filename = cal_png_14,
  width = 13,
  height = 6.7,
  units = "in",
  res = figure_dpi_14,
  bg = "white"
)
plot_calibration_panels_14()
dev.off()

tiff(
  filename = cal_tiff_14,
  width = 13,
  height = 6.7,
  units = "in",
  res = figure_dpi_14,
  compression = "lzw",
  bg = "white"
)
plot_calibration_panels_14()
dev.off()

pdf(
  file = cal_pdf_14,
  width = 13,
  height = 6.7,
  family = "sans"
)
plot_calibration_panels_14()
dev.off()

cat("14.10 Calibration图保存完成。\n")


# ==========================================================
# 14.11 保存Figure 3：DCA
# ==========================================================

DCA_png_14 <- file.path(
  figure_dir_14,
  "Figure_DCA_final_models.png"
)

DCA_tiff_14 <- file.path(
  figure_dir_14,
  "Figure_DCA_final_models.tiff"
)

DCA_pdf_14 <- file.path(
  figure_dir_14,
  "Figure_DCA_final_models.pdf"
)


png(
  filename = DCA_png_14,
  width = 9.5,
  height = 6.8,
  units = "in",
  res = figure_dpi_14,
  bg = "white"
)

plot_dca_14()

dev.off()


tiff(
  filename = DCA_tiff_14,
  width = 9.5,
  height = 6.8,
  units = "in",
  res = figure_dpi_14,
  compression = "lzw",
  bg = "white"
)

plot_dca_14()

dev.off()


pdf(
  file = DCA_pdf_14,
  width = 9.5,
  height = 6.8,
  family = "sans"
)

plot_dca_14()

dev.off()


cat(
  "14.11 DCA正文聚焦图保存完成。\n"
)


# ==========================================================
# ==========================================================
# 14.12 保存所有绘图数据与审计信息
# ==========================================================

plot_data_file_14 <- file.path(
  table_dir,
  "Figure_model_performance_plot_data.xlsx"
)

prediction_summary_14 <- data.frame(
  Model = c(
    "Preoperative model",
    "Pre-catheter-removal model"
  ),
  N = c(N_14, N_14),
  Events = c(Events_14, Events_14),
  Event_rate = c(Prevalence_14, Prevalence_14),
  Mean_predicted_risk = c(
    mean(preoperative_plot_14$prediction_mean),
    mean(precatheter_plot_14$prediction_mean)
  ),
  Min_predicted_risk = c(
    min(preoperative_plot_14$prediction_mean),
    min(precatheter_plot_14$prediction_mean)
  ),
  Max_predicted_risk = c(
    max(preoperative_plot_14$prediction_mean),
    max(precatheter_plot_14$prediction_mean)
  ),
  Calibration_axis_max = c(
    calibration_xmax_14,
    calibration_xmax_14
  ),
  stringsAsFactors = FALSE
)

# DCA关键临床阈值摘要
key_thresholds_14 <- c(
  0.05,
  0.10,
  0.15,
  0.20,
  0.25,
  0.30
)

make_key_dca_14 <- function(dca_table) {
  
  idx <- vapply(
    key_thresholds_14,
    function(x) {
      which.min(
        abs(dca_table$Threshold - x)
      )
    },
    integer(1)
  )
  
  dca_table[idx, , drop = FALSE]
}

DCA_key_thresholds_14 <- dplyr::bind_rows(
  make_key_dca_14(preoperative_dca_14),
  make_key_dca_14(precatheter_dca_14)
)

openxlsx::write.xlsx(
  list(
    ROC_AUC_summary = roc_auc_summary_14,
    ROC_AUC_by_MI = roc_auc_by_mi_14,
    ROC_Preoperative_curve = preoperative_roc_curve_14,
    ROC_Precatheter_curve = precatheter_roc_curve_14,
    Calibration_metrics = calibration_metrics_14,
    Calibration_Preoperative_groups =
      preoperative_calibration_14$groups,
    Calibration_Precatheter_groups =
      precatheter_calibration_14$groups,
    Calibration_Preoperative_curve =
      preoperative_calibration_14$curve,
    Calibration_Precatheter_curve =
      precatheter_calibration_14$curve,
    DCA_Preoperative = preoperative_dca_14,
    DCA_Precatheter = precatheter_dca_14,
    DCA_key_thresholds = DCA_key_thresholds_14,
    DCA_all_curves = DCA_all_14,
    Curve_bootstrap_QC = curve_bootstrap_qc_14,
    Prediction_summary = prediction_summary_14
  ),
  plot_data_file_14,
  overwrite = TRUE
)


# ==========================================================
# 14.13 绘图QC
# ==========================================================

auc_diff_pre_14 <- abs(
  preoperative_auc_14$summary$Apparent_AUC -
    summary_preop_13_14$Apparent_AUC
)

auc_diff_preca_14 <- abs(
  precatheter_auc_14$summary$Apparent_AUC -
    summary_precat_13_14$Apparent_AUC
)

plot_qc_14 <- data.frame(
  Item = c(
    "N",
    "Events",
    "Preoperative_AUC_difference_vs_Section13",
    "Precatheter_AUC_difference_vs_Section13",
    "Preoperative_corrected_calibration_intercept_Section13",
    "Preoperative_corrected_calibration_slope_Section13",
    "Precatheter_corrected_calibration_intercept_Section13",
    "Precatheter_corrected_calibration_slope_Section13",
    "Preoperative_max_predicted_risk",
    "Precatheter_max_predicted_risk",
    "Calibration_axis_max",
    "Preoperative_calibration_bootstrap_valid",
    "Precatheter_calibration_bootstrap_valid",
    "Preoperative_DCA_bootstrap_valid",
    "Precatheter_DCA_bootstrap_valid"
  ),
  
  Value = c(
    N_14,
    Events_14,
    auc_diff_pre_14,
    auc_diff_preca_14,
    summary_preop_13_14$Optimism_Corrected_Calibration_Intercept,
    summary_preop_13_14$Optimism_Corrected_Calibration_Slope,
    summary_precat_13_14$Optimism_Corrected_Calibration_Intercept,
    summary_precat_13_14$Optimism_Corrected_Calibration_Slope,
    max(preoperative_plot_14$prediction_mean),
    max(precatheter_plot_14$prediction_mean),
    calibration_xmax_14,
    preoperative_curve_validation_14$qc$Calibration_valid_B,
    precatheter_curve_validation_14$qc$Calibration_valid_B,
    preoperative_curve_validation_14$qc$DCA_valid_B,
    precatheter_curve_validation_14$qc$DCA_valid_B
  ),
  
  stringsAsFactors = FALSE
)


plot_qc_file_14 <- file.path(
  log_dir,
  "Figure_model_performance_QC.xlsx"
)


openxlsx::write.xlsx(
  list(
    
    Main_QC =
      plot_qc_14,
    
    Curve_bootstrap_QC =
      curve_bootstrap_qc_14,
    
    Failed_model_preop =
      data.frame(
        Bootstrap =
          preoperative_curve_validation_14$failed_model_ids
      ),
    
    Failed_model_precat =
      data.frame(
        Bootstrap =
          precatheter_curve_validation_14$failed_model_ids
      ),
    
    Failed_calib_preop =
      data.frame(
        Bootstrap =
          preoperative_curve_validation_14$failed_calibration_ids
      ),
    
    Failed_calib_precat =
      data.frame(
        Bootstrap =
          precatheter_curve_validation_14$failed_calibration_ids
      )
  ),
  
  plot_qc_file_14,
  
  overwrite = TRUE
)


if (auc_diff_pre_14 > 0.005) {
  
  warning(
    "第14部分术前AUC与第13部分相差>0.005，请核查对象是否来自同一次运行。"
  )
}


if (auc_diff_preca_14 > 0.005) {
  
  warning(
    "第14部分拔管前AUC与第13部分相差>0.005，请核查对象是否来自同一次运行。"
  )
}


if (
  max(preoperative_plot_14$prediction_mean) >
  calibration_xmax_14 ||
  max(precatheter_plot_14$prediction_mean) >
  calibration_xmax_14
) {
  
  warning(
    "Calibration坐标上限未覆盖全部患者预测概率，请提高calibration_xmax_14。"
  )
}


# ==========================================================
# 14.14 最终文件质控
# ==========================================================

required_output_files_14 <- c(
  roc_png_14,
  roc_tiff_14,
  roc_pdf_14,
  cal_png_14,
  cal_tiff_14,
  cal_pdf_14,
  DCA_png_14,
  DCA_tiff_14,
  DCA_pdf_14,
  plot_data_file_14,
  plot_qc_file_14
)


missing_output_files_14 <- required_output_files_14[
  !file.exists(
    required_output_files_14
  )
]


if (
  length(
    missing_output_files_14
  ) > 0
) {
  
  stop(
    paste0(
      "第14部分以下输出文件未成功生成：",
      paste(
        basename(
          missing_output_files_14
        ),
        collapse = ", "
      )
    )
  )
}


cat(
  "14.14 最终输出文件质控：通过。\n"
)


# ==========================================================
# 14.15 控制台输出核心结果
# ==========================================================

cat(
  "\n====================================================\n"
)

cat(
  "ROC / AUC：\n"
)

print(
  roc_auc_summary_14
)


cat(
  "\n====================================================\n"
)

cat(
  "第13部分corrected calibration intercept / slope：\n"
)

print(
  calibration_metrics_14
)


cat(
  "\n====================================================\n"
)

cat(
  "曲线级Bootstrap QC：\n"
)

print(
  curve_bootstrap_qc_14
)


cat(
  "\n====================================================\n"
)

cat(
  "Optimism-corrected DCA关键阈值：\n"
)

print(
  DCA_key_thresholds_14
)


cat(
  "\n>>> 第14部分全部完成 <<<\n"
)


cat(
  "\n请重点检查：\n",
  "1. 04_figures/Figure_ROC_final_models.png\n",
  "2. 04_figures/Figure_Calibration_final_models.png\n",
  "3. 04_figures/Figure_DCA_final_models.png\n",
  "4. 02_tables/Figure_model_performance_plot_data.xlsx\n",
  "5. 05_logs/Figure_model_performance_QC.xlsx\n",
  sep = ""
)

############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
############################################################


############################################################
############################################################
############################################################
# 15. Final-model nomograms
#
# Uses the FINAL MI-pooled logistic coefficients from Section 12.
# Does NOT refit models and does NOT apply post-hoc shrinkage.
############################################################

cat("\n>>> Section 15 started: Final-model nomograms <<<\n")
flush.console()

# ----------------------------------------------------------
# 15.1 Basic checks
# ----------------------------------------------------------

need_obj_15 <- c("table_dir", "model_dir", "log_dir")
miss_obj_15 <- need_obj_15[!vapply(need_obj_15, exists, logical(1), inherits = TRUE)]
if (length(miss_obj_15) > 0) {
  stop(paste0("Missing objects: ", paste(miss_obj_15, collapse = ", ")))
}

if (!requireNamespace("openxlsx", quietly = TRUE)) {
  stop("Package 'openxlsx' is required.")
}

figure_dir_15 <- file.path(dirname(table_dir), "04_figures")
dir.create(figure_dir_15, recursive = TRUE, showWarnings = FALSE)

# Load Section 12 model objects if needed.
if (!exists("preoperative_final_model_12", inherits = TRUE) ||
    !exists("precatheter_final_model_12", inherits = TRUE)) {
  
  rds_15 <- file.path(model_dir, "Final_multivariable_model_objects.rds")
  if (!file.exists(rds_15)) {
    stop(paste0("Cannot find final model objects: ", rds_15))
  }
  
  tmp_15 <- readRDS(rds_15)
  preoperative_final_model_12 <- tmp_15$Preoperative_final
  precatheter_final_model_12 <- tmp_15$Precatheter_final
}

cat("15.1 Basic checks passed.\n")

# ----------------------------------------------------------
# 15.2 Helper functions
# ----------------------------------------------------------

get_fits_15 <- function(model_object) {
  if (!is.null(model_object$fits)) return(model_object$fits)
  if (!is.null(model_object$fit_list)) return(model_object$fit_list)
  stop("No fits/fit_list found in final model object.")
}

get_frame_15 <- function(model_object) {
  ff <- get_fits_15(model_object)
  if (length(ff) < 1 || is.null(ff[[1]]$model)) {
    stop("Cannot retrieve model frame from final model object.")
  }
  ff[[1]]$model
}

binary01_15 <- function(x, varname = "binary variable") {
  z <- trimws(tolower(as.character(x)))
  out <- rep(NA_real_, length(z))
  out[z %in% c("0", "no", "n", "false")] <- 0
  out[z %in% c("1", "yes", "y", "true")] <- 1
  if (anyNA(out)) {
    bad <- unique(as.character(x)[is.na(out)])
    stop(paste0(
      "Cannot map ", varname, " to 0/1. Values: ",
      paste(bad, collapse = ", ")
    ))
  }
  out
}

extract_coef_15 <- function(model_object, model_name) {
  pt <- model_object$pooled_table
  if (is.null(pt) || !all(c("Variable", "Beta") %in% names(pt))) {
    stop(paste0(model_name, ": pooled_table with Variable/Beta not found."))
  }
  
  ii <- which(pt$Variable == "(Intercept)")
  if (length(ii) != 1) stop(paste0(model_name, ": intercept not unique."))
  
  pp <- pt[pt$Variable != "(Intercept)", c("Variable", "Beta"), drop = FALSE]
  if (nrow(pp) < 1) stop(paste0(model_name, ": no predictors."))
  
  list(
    intercept = as.numeric(pt$Beta[ii]),
    beta = stats::setNames(as.numeric(pp$Beta), pp$Variable)
  )
}

pretty_inside_15 <- function(rng, n = 6) {
  z <- pretty(rng, n = n)
  z <- z[z >= min(rng) & z <= max(rng)]
  z <- unique(c(min(rng), z, max(rng)))
  sort(z)
}

format_num_15 <- function(x) {
  if (all(abs(x - round(x)) < 1e-8)) {
    format(round(x), trim = TRUE, scientific = FALSE)
  } else {
    format(round(x, 1), trim = TRUE, scientific = FALSE)
  }
}

# ----------------------------------------------------------
# 15.3 Explicit variable definitions
# ----------------------------------------------------------

binary_vars_15 <- c(
  "incomplete_emptying",
  "detrusor_thickening",
  "concomitant_hysterectomy"
)

continuous_vars_15 <- c(
  "bmi",
  "blood_loss_ml"
)

label_map_15 <- c(
  bmi = "BMI (kg/m²)",
  incomplete_emptying = "Incomplete emptying",
  detrusor_thickening = "Detrusor thickening",
  concomitant_hysterectomy = "Concomitant hysterectomy",
  blood_loss_ml = "Estimated blood loss (mL)"
)

# ----------------------------------------------------------
# 15.4 Build one nomogram specification
# ----------------------------------------------------------

build_spec_15 <- function(model_object, model_name, panel_letter) {
  
  cf <- extract_coef_15(model_object, model_name)
  mf <- get_frame_15(model_object)
  vars <- names(cf$beta)
  
  missing_frame <- setdiff(vars, names(mf))
  if (length(missing_frame) > 0) {
    stop(paste0(model_name, ": variables missing in model frame: ",
                paste(missing_frame, collapse = ", ")))
  }
  
  unknown_vars <- setdiff(vars, c(binary_vars_15, continuous_vars_15))
  if (length(unknown_vars) > 0) {
    stop(paste0(model_name, ": variable type not defined: ",
                paste(unknown_vars, collapse = ", ")))
  }
  
  info <- vector("list", length(vars))
  names(info) <- vars
  
  for (v in vars) {
    b <- unname(cf$beta[v])
    
    if (v %in% binary_vars_15) {
      xx <- binary01_15(mf[[v]], v)
      rng <- c(0, 1)
      ticks <- c(0, 1)
      tick_labels <- c("No", "Yes")
    } else {
      xx <- suppressWarnings(as.numeric(as.character(mf[[v]])))
      if (anyNA(xx)) stop(paste0(model_name, ": non-numeric values in ", v))
      rng <- range(xx, na.rm = TRUE)
      ticks <- pretty_inside_15(rng, n = 5)
      tick_labels <- format_num_15(ticks)
    }
    
    contrib_rng <- b * rng
    min_contrib <- min(contrib_rng)
    max_contrib <- max(contrib_rng)
    effect_range <- max_contrib - min_contrib
    
    info[[v]] <- list(
      variable = v,
      label = if (v %in% names(label_map_15)) unname(label_map_15[v]) else v,
      beta = b,
      range = rng,
      ticks = ticks,
      tick_labels = tick_labels,
      min_contrib = min_contrib,
      max_contrib = max_contrib,
      effect_range = effect_range
    )
  }
  
  max_effect <- max(vapply(info, function(z) z$effect_range, numeric(1)))
  if (!is.finite(max_effect) || max_effect <= 0) {
    stop(paste0(model_name, ": invalid predictor effect range."))
  }
  
  point_scale <- 100 / max_effect
  
  for (v in vars) {
    z <- info[[v]]
    z$tick_points <- (z$beta * z$ticks - z$min_contrib) * point_scale
    z$max_points <- z$effect_range * point_scale
    info[[v]] <- z
  }
  
  min_sum <- sum(vapply(info, function(z) z$min_contrib, numeric(1)))
  total_max <- sum(vapply(info, function(z) z$max_points, numeric(1)))
  
  risk_from_points <- function(total_points) {
    stats::plogis(cf$intercept + min_sum + total_points / point_scale)
  }
  
  points_from_risk <- function(risk) {
    (stats::qlogis(risk) - cf$intercept - min_sum) * point_scale
  }
  
  list(
    model_name = model_name,
    panel_letter = panel_letter,
    intercept = cf$intercept,
    beta = cf$beta,
    model_frame = mf,
    variables = vars,
    info = info,
    point_scale = point_scale,
    min_sum = min_sum,
    total_max = total_max,
    risk_from_points = risk_from_points,
    points_from_risk = points_from_risk
  )
}

preop_nom_15 <- build_spec_15(
  preoperative_final_model_12,
  "Preoperative model",
  "A"
)

precat_nom_15 <- build_spec_15(
  precatheter_final_model_12,
  "Pre-catheter-removal model",
  "B"
)

cat("15.4 Nomogram specifications built successfully.\n")

# ----------------------------------------------------------
# 15.5 Mathematical QC
# ----------------------------------------------------------

check_points_qc_15 <- function(spec) {
  mf <- spec$model_frame
  n <- nrow(mf)
  lp_direct <- rep(spec$intercept, n)
  total_points <- rep(0, n)
  
  for (v in spec$variables) {
    b <- unname(spec$beta[v])
    z <- spec$info[[v]]
    
    if (v %in% binary_vars_15) {
      x <- binary01_15(mf[[v]], v)
    } else {
      x <- as.numeric(as.character(mf[[v]]))
    }
    
    lp_direct <- lp_direct + b * x
    total_points <- total_points + (b * x - z$min_contrib) * spec$point_scale
  }
  
  p_direct <- stats::plogis(lp_direct)
  p_points <- spec$risk_from_points(total_points)
  
  data.frame(
    Model = spec$model_name,
    N = n,
    Max_absolute_probability_difference = max(abs(p_direct - p_points), na.rm = TRUE),
    stringsAsFactors = FALSE
  )
}

qc_15 <- rbind(
  check_points_qc_15(preop_nom_15),
  check_points_qc_15(precat_nom_15)
)

if (any(qc_15$Max_absolute_probability_difference > 1e-10)) {
  warning("Nomogram point-to-risk mapping differs from original logistic equation by >1e-10.")
}

cat("15.5 Mathematical QC completed.\n")

# ----------------------------------------------------------
# 15.6 Publication-style nomogram plotting
#
# Design principles:
# - Preserve the Section 12 MI-pooled coefficients exactly.
# - Use a wide, single-page clinical nomogram layout.
# - Reserve a fixed left column for readable predictor labels.
# - Use regular, uncluttered Total Points and risk axes.
# - Single-model figures are the primary outputs.
# - The combined figure stacks A/B vertically rather than side-by-side.
# ----------------------------------------------------------

# Clinical display labels.
label_map_15 <- c(
  bmi = "BMI (kg/m²)",
  incomplete_emptying = "Incomplete emptying",
  detrusor_thickening = "Detrusor thickening",
  concomitant_hysterectomy = "Concomitant hysterectomy",
  blood_loss_ml = "Estimated blood loss (mL)"
)

# Explicit display ticks. These change appearance only, not model mathematics.
# BMI is handled dynamically so BOTH actual endpoints are always shown.
# This prevents the right end of the BMI axis from appearing "open".
tick_override_15 <- list(
  blood_loss_ml = c(5, 50, 100, 150, 200, 250, 300, 330)
)

format_display_tick_15 <- function(v, x) {
  if (v == "bmi") {
    out <- ifelse(
      abs(x - round(x)) < 0.05,
      as.character(round(x)),
      format(round(x, 1), trim = TRUE, scientific = FALSE)
    )
    return(out)
  }
  
  if (v == "blood_loss_ml") {
    return(as.character(round(x)))
  }
  
  as.character(x)
}

# Return the EXACT ticks used for plotting/export.
# For BMI, the observed minimum and maximum are forced onto the axis.
# Example: if the observed minimum is 16.53, it is plotted at the true
# endpoint and labelled 16.5. No extrapolation is introduced.
get_display_ticks_15 <- function(spec, v) {
  
  z <- spec$info[[v]]
  
  if (v %in% binary_vars_15) {
    tick_values <- c(0, 1)
    tick_labels <- c("No", "Yes")
    
  } else if (v == "bmi") {
    
    bmi_min <- min(z$range, na.rm = TRUE)
    bmi_max <- max(z$range, na.rm = TRUE)
    
    interior <- c(20, 25, 30)
    interior <- interior[
      interior > bmi_min + 1e-8 &
        interior < bmi_max - 1e-8
    ]
    
    # Descending clinical order: high BMI on the left, low BMI on the right.
    tick_values <- sort(
      unique(c(bmi_max, interior, bmi_min)),
      decreasing = TRUE
    )
    
    tick_labels <- format_display_tick_15(v, tick_values)
    
  } else if (v %in% names(tick_override_15)) {
    
    tick_values <- tick_override_15[[v]]
    tick_values <- tick_values[
      tick_values >= min(z$range) - 1e-8 &
        tick_values <= max(z$range) + 1e-8
    ]
    
    # Always include exact observed endpoints for a visually closed axis.
    tick_values <- sort(
      unique(c(min(z$range), tick_values, max(z$range)))
    )
    
    tick_labels <- format_display_tick_15(v, tick_values)
    
  } else {
    
    tick_values <- z$ticks
    tick_labels <- format_display_tick_15(v, tick_values)
  }
  
  tick_points <- (
    z$beta * tick_values -
      z$min_contrib
  ) * spec$point_scale
  
  list(
    values = tick_values,
    labels = tick_labels,
    points = tick_points
  )
}

nice_total_axis_15 <- function(total_max) {
  if (total_max <= 180) {
    step <- 20
  } else if (total_max <= 300) {
    step <- 40
  } else if (total_max <= 450) {
    step <- 50
  } else {
    step <- 100
  }
  axis_max <- ceiling(total_max / step) * step
  list(
    max = axis_max,
    ticks = seq(0, axis_max, by = step)
  )
}

# Risk marks used in the published nomogram.
risk_candidates_15 <- c(
  0.01, 0.02, 0.05, 0.10,
  0.20, 0.30, 0.50, 0.70
)

# ----------------------------------------------------------
# 15.6A Draw one wide nomogram
# ----------------------------------------------------------

plot_nomogram_wide_15 <- function(
    spec,
    show_header = TRUE,
    panel_letter = NULL,
    title_text = NULL
) {
  
  vars <- spec$variables
  nvar <- length(vars)
  
  # Drawing region. Labels live inside the plotting region, so clipping is avoided.
  x_label <- 0.035
  x_axis0 <- 0.305
  x_axis1 <- 0.965
  axis_width <- x_axis1 - x_axis0
  
  # Vertical positions.
  y_title <- 0.965
  y_points <- 0.865
  
  if (nvar <= 3) {
    y_vars <- seq(0.70, by = -0.12, length.out = nvar)
    y_total <- min(y_vars) - 0.18
    y_risk <- y_total - 0.15
  } else {
    y_vars <- seq(0.72, by = -0.095, length.out = nvar)
    y_total <- min(y_vars) - 0.14
    y_risk <- y_total - 0.13
  }
  
  plot.new()
  plot.window(
    xlim = c(0, 1),
    ylim = c(0, 1),
    xaxs = "i",
    yaxs = "i"
  )
  
  # Header for the stacked A/B figure.
  if (show_header) {
    if (!is.null(panel_letter)) {
      text(
        0.015, y_title,
        labels = panel_letter,
        adj = c(0, 1),
        cex = 1.45,
        font = 2
      )
    }
    
    if (is.null(title_text)) {
      title_text <- spec$model_name
    }
    
    text(
      0.50, y_title,
      labels = title_text,
      adj = c(0.5, 1),
      cex = 1.24,
      font = 2
    )
  }
  
  # --------------------------------------------------------
  # Points axis: 0-100, with minor ticks.
  # --------------------------------------------------------
  
  text(
    x_label, y_points,
    labels = "Points",
    adj = c(0, 0.5),
    cex = 1.05
  )
  
  segments(
    x_axis0, y_points,
    x_axis1, y_points,
    lwd = 1.05
  )
  
  main_points <- seq(0, 100, by = 10)
  minor_points <- seq(0, 100, by = 2)
  
  x_minor <- x_axis0 + minor_points / 100 * axis_width
  segments(
    x_minor, y_points,
    x_minor, y_points - 0.008,
    lwd = 0.55
  )
  
  x_main <- x_axis0 + main_points / 100 * axis_width
  segments(
    x_main, y_points,
    x_main, y_points - 0.015,
    lwd = 0.9
  )
  
  text(
    x_main, y_points + 0.027,
    labels = main_points,
    adj = c(0.5, 0),
    cex = 0.78
  )
  
  # --------------------------------------------------------
  # Predictor axes.
  # --------------------------------------------------------
  
  for (k in seq_along(vars)) {
    v <- vars[k]
    z <- spec$info[[v]]
    yy <- y_vars[k]
    
    label_now <- if (v %in% names(label_map_15)) {
      unname(label_map_15[v])
    } else {
      z$label
    }
    
    text(
      x_label, yy,
      labels = label_now,
      adj = c(0, 0.5),
      cex = 1.02
    )
    
    # Axis length is proportional to the maximum number of points contributed by the predictor.
    x_end <- x_axis0 + (z$max_points / 100) * axis_width
    
    segments(
      x_axis0, yy,
      x_end, yy,
      lwd = 1.05
    )
    
    display_ticks <- get_display_ticks_15(
      spec = spec,
      v = v
    )
    
    tick_values <- display_ticks$values
    tick_labels <- display_ticks$labels
    tick_points <- display_ticks$points
    
    x_ticks <- x_axis0 + (tick_points / 100) * axis_width
    
    segments(
      x_ticks, yy,
      x_ticks, yy - 0.015,
      lwd = 0.90
    )
    
    text(
      x_ticks, yy - 0.027,
      labels = tick_labels,
      adj = c(0.5, 1),
      cex = 0.79
    )
  }
  
  # --------------------------------------------------------
  # Total Points axis.
  # --------------------------------------------------------
  
  total_axis <- nice_total_axis_15(spec$total_max)
  total_axis_max <- total_axis$max
  total_ticks <- total_axis$ticks
  
  text(
    x_label, y_total,
    labels = "Total Points",
    adj = c(0, 0.5),
    cex = 1.05
  )
  
  segments(
    x_axis0, y_total,
    x_axis1, y_total,
    lwd = 1.10
  )
  
  # Minor total-points ticks.
  minor_step <- if (total_axis_max <= 300) 10 else 25
  total_minor <- seq(0, total_axis_max, by = minor_step)
  x_total_minor <- x_axis0 + total_minor / total_axis_max * axis_width
  
  segments(
    x_total_minor, y_total,
    x_total_minor, y_total - 0.008,
    lwd = 0.55
  )
  
  x_total <- x_axis0 + total_ticks / total_axis_max * axis_width
  
  segments(
    x_total, y_total,
    x_total, y_total - 0.015,
    lwd = 0.90
  )
  
  text(
    x_total, y_total - 0.027,
    labels = total_ticks,
    adj = c(0.5, 1),
    cex = 0.79
  )
  
  # --------------------------------------------------------
  # Predicted POUR probability axis.
  # --------------------------------------------------------
  
  risk_points <- spec$points_from_risk(risk_candidates_15)
  
  keep_risk <- is.finite(risk_points) &
    risk_points >= 0 &
    risk_points <= spec$total_max
  
  risk_ticks <- risk_candidates_15[keep_risk]
  risk_points <- risk_points[keep_risk]
  
  x_risk <- x_axis0 + risk_points / total_axis_max * axis_width
  
  text(
    x_label, y_risk,
    labels = "Predicted probability of POUR",
    adj = c(0, 0.5),
    cex = 1.05
  )
  
  if (length(x_risk) >= 2) {
    segments(
      min(x_risk), y_risk,
      max(x_risk), y_risk,
      lwd = 1.10
    )
  }
  
  segments(
    x_risk, y_risk,
    x_risk, y_risk - 0.015,
    lwd = 0.90
  )
  
  risk_labels <- paste0(round(100 * risk_ticks), "%")
  
  text(
    x_risk, y_risk - 0.027,
    labels = risk_labels,
    adj = c(0.5, 1),
    cex = 0.78
  )
  
  invisible(
    list(
      total_axis_max = total_axis_max,
      total_ticks = total_ticks,
      risk_ticks = risk_ticks,
      risk_points = risk_points
    )
  )
}

# ----------------------------------------------------------
# 15.7 Save publication-style figures
#
# FINAL OUTPUT POLICY:
# - TIFF only for figures
# - 600 dpi
# - LZW compression
# - Keep two single-model figures plus one stacked A/B figure
# ----------------------------------------------------------

figure_dpi_15 <- 600

preop_tiff_15 <- file.path(
  figure_dir_15,
  "Figure_Nomogram_Preoperative.tiff"
)

precat_tiff_15 <- file.path(
  figure_dir_15,
  "Figure_Nomogram_Precatheter.tiff"
)

combined_tiff_15 <- file.path(
  figure_dir_15,
  "Figure_Nomogram_final_models.tiff"
)

save_single_nomogram_tiff_15 <- function(
    spec,
    tiff_file,
    height_in
) {
  
  tiff(
    filename = tiff_file,
    width = 13.0,
    height = height_in,
    units = "in",
    res = figure_dpi_15,
    compression = "lzw",
    bg = "white"
  )
  
  par(
    mar = c(0.6, 0.6, 0.6, 0.6),
    family = "sans"
  )
  
  plot_nomogram_wide_15(
    spec = spec,
    show_header = FALSE
  )
  
  dev.off()
}

save_single_nomogram_tiff_15(
  spec = preop_nom_15,
  tiff_file = preop_tiff_15,
  height_in = 6.8
)

save_single_nomogram_tiff_15(
  spec = precat_nom_15,
  tiff_file = precat_tiff_15,
  height_in = 7.8
)

# Combined A/B figure: vertical stacking preserves full axis width.
tiff(
  filename = combined_tiff_15,
  width = 13.0,
  height = 13.8,
  units = "in",
  res = figure_dpi_15,
  compression = "lzw",
  bg = "white"
)

par(
  mfrow = c(2, 1),
  mar = c(0.4, 0.4, 0.4, 0.4),
  oma = c(0, 0, 0, 0),
  family = "sans"
)

plot_nomogram_wide_15(
  spec = preop_nom_15,
  show_header = TRUE,
  panel_letter = "A",
  title_text = "Preoperative model"
)

plot_nomogram_wide_15(
  spec = precat_nom_15,
  show_header = TRUE,
  panel_letter = "B",
  title_text = "Pre-catheter-removal model"
)

dev.off()

cat("15.7 TIFF nomogram figures saved.\n")

# ----------------------------------------------------------
# 15.8 Export scoring tables
# ----------------------------------------------------------

make_coef_table_15 <- function(spec) {
  rows <- list(
    data.frame(
      Model = spec$model_name,
      Variable = "(Intercept)",
      Beta = spec$intercept,
      Type = "Intercept",
      Display_min = NA_real_,
      Display_max = NA_real_,
      Maximum_points = NA_real_,
      stringsAsFactors = FALSE
    )
  )
  
  for (v in spec$variables) {
    z <- spec$info[[v]]
    rows[[length(rows) + 1]] <- data.frame(
      Model = spec$model_name,
      Variable = v,
      Beta = z$beta,
      Type = if (v %in% binary_vars_15) "Binary" else "Continuous",
      Display_min = z$range[1],
      Display_max = z$range[2],
      Maximum_points = z$max_points,
      stringsAsFactors = FALSE
    )
  }
  
  do.call(rbind, rows)
}

make_tick_table_15 <- function(spec) {
  out <- list()
  
  for (v in spec$variables) {
    
    z <- spec$info[[v]]
    display_ticks <- get_display_ticks_15(
      spec = spec,
      v = v
    )
    
    out[[length(out) + 1]] <- data.frame(
      Model = spec$model_name,
      Variable = v,
      Display_label = if (v %in% names(label_map_15)) {
        unname(label_map_15[v])
      } else {
        z$label
      },
      Tick_value = display_ticks$values,
      Tick_label = display_ticks$labels,
      Points = display_ticks$points,
      stringsAsFactors = FALSE
    )
  }
  
  do.call(rbind, out)
}

make_risk_table_15 <- function(spec) {
  tp <- seq(0, spec$total_max, length.out = 101)
  data.frame(
    Model = spec$model_name,
    Total_points = tp,
    Predicted_POUR_risk = spec$risk_from_points(tp),
    stringsAsFactors = FALSE
  )
}

coef_table_15 <- rbind(make_coef_table_15(preop_nom_15), make_coef_table_15(precat_nom_15))
tick_table_15 <- rbind(make_tick_table_15(preop_nom_15), make_tick_table_15(precat_nom_15))
risk_table_15 <- rbind(make_risk_table_15(preop_nom_15), make_risk_table_15(precat_nom_15))

xlsx_15 <- file.path(table_dir, "Nomogram_scoring_and_risk_mapping.xlsx")
openxlsx::write.xlsx(
  list(
    Coefficients = coef_table_15,
    Variable_ticks = tick_table_15,
    Points_to_risk = risk_table_15,
    QC = qc_15
  ),
  xlsx_15,
  overwrite = TRUE
)

qc_file_15 <- file.path(log_dir, "Nomogram_QC.xlsx")
openxlsx::write.xlsx(list(QC = qc_15), qc_file_15, overwrite = TRUE)

rds_nom_15 <- file.path(model_dir, "Final_nomogram_objects.rds")
saveRDS(list(Preoperative = preop_nom_15, Precatheter = precat_nom_15), rds_nom_15)

# ----------------------------------------------------------
# 15.9 Final file QC and console output
# ----------------------------------------------------------

required_files_15 <- c(
  preop_tiff_15,
  precat_tiff_15,
  combined_tiff_15,
  xlsx_15,
  qc_file_15,
  rds_nom_15
)

missing_files_15 <- required_files_15[!file.exists(required_files_15)]
if (length(missing_files_15) > 0) {
  stop(paste0("Section 15 output missing: ", paste(basename(missing_files_15), collapse = ", ")))
}

cat("\nNomogram QC:\n")
print(qc_15)
cat("\n>>> Section 15 completed successfully <<<\n")
cat("Please inspect:\n")
cat("1. 04_figures/Figure_Nomogram_Preoperative.tiff\n")
cat("2. 04_figures/Figure_Nomogram_Precatheter.tiff\n")
cat("3. 04_figures/Figure_Nomogram_final_models.tiff\n")
cat("4. 02_tables/Nomogram_scoring_and_risk_mapping.xlsx\n")
cat("5. 05_logs/Nomogram_QC.xlsx\n")


############################################################
############################################################
############################################################
############################################################
############################################################
############################################################
# 16. Penalized Logistic Regression Sensitivity Analysis
#
# Purpose:
# 1. Use the same candidate-variable sets as the traditional
#    multivariable Logistic route:
#       - Preoperative full candidate model: 8 predictors
#       - Pre-catheter-removal full candidate model: 11 predictors
# 2. Fit three penalized Logistic approaches in each of the
#    20 multiply imputed datasets:
#       - Ridge       : alpha = 0
#       - Elastic Net : alpha = 0.5
#       - LASSO       : alpha = 1
# 3. Use the SAME stratified 5-fold assignments:
#       - across all 20 imputations
#       - across Ridge / Elastic Net / LASSO
#       - across the two clinical time points
# 4. Select lambda by cross-validated binomial deviance.
# 5. Report BOTH:
#       - lambda.min
#       - lambda.1se
# 6. Use prevalidated (out-of-fold) predictions returned by
#    cv.glmnet(keep = TRUE) to calculate:
#       - CV AUC
#       - CV Brier score (conventional binary scale)
#       - CV binomial deviance
# 7. Penalized coefficients are NOT Rubin-pooled.
#    Instead, coefficient direction/magnitude and selection
#    frequency are summarized across 20 imputations.
# 8. Ridge is a shrinkage method, not a variable-selection
#    method; therefore selection frequency is not interpreted
#    for Ridge.
#
# IMPORTANT:
# This section is a SENSITIVITY ANALYSIS.
# It does NOT replace the primary traditional Logistic models,
# their 1000-bootstrap internal validation, ROC, calibration,
# DCA, or nomograms.
############################################################


cat("\n>>> Section 16 started: Penalized-regression sensitivity analysis <<<\n")
flush.console()


# ==========================================================
# 16.1 Required packages and objects
# ==========================================================

required_packages_16 <- c(
  "glmnet",
  "mice",
  "dplyr",
  "openxlsx"
)

missing_packages_16 <- required_packages_16[
  !vapply(
    required_packages_16,
    requireNamespace,
    logical(1),
    quietly = TRUE
  )
]

if (length(missing_packages_16) > 0) {
  stop(
    paste0(
      "Section 16 is missing required R packages: ",
      paste(missing_packages_16, collapse = ", ")
    )
  )
}


required_objects_16 <- c(
  "imp_preoperative",
  "imp_precatheter",
  "preoperative_selected_model",
  "precatheter_selected_model",
  "outcome_var",
  "table_dir",
  "model_dir",
  "log_dir"
)

missing_objects_16 <- required_objects_16[
  !vapply(
    required_objects_16,
    exists,
    logical(1),
    inherits = TRUE
  )
]

if (length(missing_objects_16) > 0) {
  stop(
    paste0(
      "Section 16 is missing prerequisite objects: ",
      paste(missing_objects_16, collapse = ", ")
    )
  )
}


if (imp_preoperative$m != 20) {
  stop("The preoperative MICE object does not contain 20 imputations.")
}

if (imp_precatheter$m != 20) {
  stop("The pre-catheter MICE object does not contain 20 imputations.")
}


dir.create(
  table_dir,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  model_dir,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  log_dir,
  recursive = TRUE,
  showWarnings = FALSE
)


cat("16.1 Package/object checks: passed.\n")


# ==========================================================
# 16.2 Recover the locked candidate-variable sets
#
# We keep popq_bp and remove posterior_wall_stage_model,
# exactly as in Sections 11-12.
# ==========================================================

preoperative_full_vars_16 <- setdiff(
  preoperative_selected_model,
  "posterior_wall_stage_model"
)

precatheter_full_vars_16 <- setdiff(
  precatheter_selected_model,
  "posterior_wall_stage_model"
)


if (!"popq_bp" %in% preoperative_full_vars_16) {
  stop("Preoperative penalized candidate set is missing popq_bp.")
}

if (!"popq_bp" %in% precatheter_full_vars_16) {
  stop("Pre-catheter penalized candidate set is missing popq_bp.")
}

if ("posterior_wall_stage_model" %in% preoperative_full_vars_16) {
  stop("posterior_wall_stage_model remains in the preoperative candidate set.")
}

if ("posterior_wall_stage_model" %in% precatheter_full_vars_16) {
  stop("posterior_wall_stage_model remains in the pre-catheter candidate set.")
}


if (length(preoperative_full_vars_16) != 8) {
  warning(
    paste0(
      "Preoperative candidate set contains ",
      length(preoperative_full_vars_16),
      " predictors rather than the expected 8."
    )
  )
}

if (length(precatheter_full_vars_16) != 11) {
  warning(
    paste0(
      "Pre-catheter candidate set contains ",
      length(precatheter_full_vars_16),
      " predictors rather than the expected 11."
    )
  )
}


candidate_table_16 <- dplyr::bind_rows(
  data.frame(
    Timepoint = "Preoperative",
    Variable = preoperative_full_vars_16,
    stringsAsFactors = FALSE
  ),
  data.frame(
    Timepoint = "Pre-catheter-removal",
    Variable = precatheter_full_vars_16,
    stringsAsFactors = FALSE
  )
)


cat("\nPreoperative penalized candidate predictors:\n")
print(preoperative_full_vars_16)

cat("\nPre-catheter-removal penalized candidate predictors:\n")
print(precatheter_full_vars_16)

cat("16.2 Candidate-variable sets: locked.\n")


# ==========================================================
# 16.3 Recover primary Logistic-model predictors
#
# These are only used to flag whether a penalized-regression
# predictor was also retained in the primary Logistic model.
# ==========================================================

primary_preop_vars_16 <- character()
primary_precat_vars_16 <- character()


if (
  exists(
    "preoperative_final_model_12",
    inherits = TRUE
  ) &&
  !is.null(
    preoperative_final_model_12$predictor_vars
  )
) {
  primary_preop_vars_16 <- preoperative_final_model_12$predictor_vars
}

if (
  exists(
    "precatheter_final_model_12",
    inherits = TRUE
  ) &&
  !is.null(
    precatheter_final_model_12$predictor_vars
  )
) {
  primary_precat_vars_16 <- precatheter_final_model_12$predictor_vars
}


# Try loading Section 12 model objects if needed
if (
  length(primary_preop_vars_16) == 0 ||
  length(primary_precat_vars_16) == 0
) {
  
  final_model_rds_16 <- file.path(
    model_dir,
    "Final_multivariable_model_objects.rds"
  )
  
  if (file.exists(final_model_rds_16)) {
    
    final_models_loaded_16 <- readRDS(
      final_model_rds_16
    )
    
    if (
      length(primary_preop_vars_16) == 0 &&
      !is.null(
        final_models_loaded_16$Preoperative_final$predictor_vars
      )
    ) {
      primary_preop_vars_16 <-
        final_models_loaded_16$Preoperative_final$predictor_vars
    }
    
    if (
      length(primary_precat_vars_16) == 0 &&
      !is.null(
        final_models_loaded_16$Precatheter_final$predictor_vars
      )
    ) {
      primary_precat_vars_16 <-
        final_models_loaded_16$Precatheter_final$predictor_vars
    }
  }
}


cat("\nPrimary Logistic predictors used for comparison:\n")
cat(
  "  Preoperative: ",
  paste(primary_preop_vars_16, collapse = ", "),
  "\n",
  sep = ""
)
cat(
  "  Pre-catheter-removal: ",
  paste(primary_precat_vars_16, collapse = ", "),
  "\n",
  sep = ""
)


# ==========================================================
# 16.4 Helper functions
# ==========================================================

# ----------------------------------------------------------
# 16.4A Safe mean / SD
# ----------------------------------------------------------

safe_mean_16 <- function(x) {
  if (length(x) == 0 || all(is.na(x))) {
    return(NA_real_)
  }
  
  mean(
    x,
    na.rm = TRUE
  )
}


safe_sd_16 <- function(x) {
  valid_n <- sum(
    is.finite(x)
  )
  
  if (valid_n <= 1) {
    return(NA_real_)
  }
  
  stats::sd(
    x,
    na.rm = TRUE
  )
}


safe_median_16 <- function(x) {
  if (length(x) == 0 || all(is.na(x))) {
    return(NA_real_)
  }
  
  stats::median(
    x,
    na.rm = TRUE
  )
}


# ----------------------------------------------------------
# 16.4B Outcome -> explicit 0/1
# ----------------------------------------------------------

coerce_binary_outcome_16 <- function(
    x,
    variable_name = "Outcome"
) {
  
  if (is.logical(x)) {
    x <- as.integer(x)
  }
  
  if (
    is.factor(x) ||
    is.character(x)
  ) {
    
    z <- trimws(
      as.character(x)
    )
    
    values_now <- unique(
      z[
        !is.na(z)
      ]
    )
    
    if (
      !all(
        values_now %in%
        c("0", "1")
      )
    ) {
      stop(
        paste0(
          variable_name,
          " is not explicitly coded as 0/1. Values: ",
          paste(values_now, collapse = ", ")
        )
      )
    }
    
    x <- as.numeric(z)
  }
  
  x <- as.numeric(x)
  
  if (anyNA(x)) {
    stop(
      paste0(
        variable_name,
        " contains NA after conversion to 0/1."
      )
    )
  }
  
  if (!all(x %in% c(0, 1))) {
    stop(
      paste0(
        variable_name,
        " is not a binary 0/1 outcome."
      )
    )
  }
  
  x
}


# ----------------------------------------------------------
# 16.4C Build a glmnet design matrix
#
# model.matrix() automatically converts two-level factors
# to one indicator term under the usual treatment contrasts.
# Current Section 12 route requires all candidate variables
# to contribute exactly 1 df.
# ----------------------------------------------------------

build_design_matrix_16 <- function(
    data_now,
    predictor_vars,
    expected_columns = NULL
) {
  
  formula_now <- stats::reformulate(
    predictor_vars
  )
  
  x_now <- stats::model.matrix(
    formula_now,
    data = data_now
  )
  
  if ("(Intercept)" %in% colnames(x_now)) {
    x_now <- x_now[
      ,
      colnames(x_now) != "(Intercept)",
      drop = FALSE
    ]
  }
  
  storage.mode(x_now) <- "double"
  
  
  if (anyNA(x_now)) {
    stop(
      "The glmnet design matrix contains missing values."
    )
  }
  
  if (
    ncol(x_now) !=
    length(predictor_vars)
  ) {
    stop(
      paste0(
        "The penalized model produced ",
        ncol(x_now),
        " design-matrix columns from ",
        length(predictor_vars),
        " candidate predictors. ",
        "Section 16 currently assumes all predictors are 1-df variables."
      )
    )
  }
  
  if (!is.null(expected_columns)) {
    
    if (
      !identical(
        colnames(x_now),
        expected_columns
      )
    ) {
      stop(
        paste0(
          "Design-matrix columns differ across imputations. ",
          "Expected: ",
          paste(expected_columns, collapse = ", "),
          "; observed: ",
          paste(colnames(x_now), collapse = ", ")
        )
      )
    }
  }
  
  x_now
}


# ----------------------------------------------------------
# 16.4D Map original variables to model-matrix coefficient terms
# ----------------------------------------------------------

make_term_map_16 <- function(
    data_now,
    predictor_vars
) {
  
  formula_now <- stats::reformulate(
    predictor_vars
  )
  
  mm <- stats::model.matrix(
    formula_now,
    data = data_now
  )
  
  assign_vec <- attr(
    mm,
    "assign"
  )
  
  term_labels <- attr(
    stats::terms(formula_now),
    "term.labels"
  )
  
  result_list <- vector(
    "list",
    length(term_labels)
  )
  
  for (i in seq_along(term_labels)) {
    
    cols_now <- colnames(mm)[
      assign_vec == i
    ]
    
    if (length(cols_now) != 1) {
      stop(
        paste0(
          "Variable ",
          term_labels[i],
          " generated ",
          length(cols_now),
          " coefficient columns. ",
          "Section 16 requires 1-df predictors."
        )
      )
    }
    
    result_list[[i]] <- data.frame(
      Variable = term_labels[i],
      Term = cols_now,
      stringsAsFactors = FALSE
    )
  }
  
  dplyr::bind_rows(
    result_list
  )
}


# ----------------------------------------------------------
# 16.4E Create one stratified 5-fold split
#
# The same patient-level fold allocation is reused across:
# - 20 imputations
# - 3 penalty methods
# - both clinical time points
# ----------------------------------------------------------

make_stratified_foldid_16 <- function(
    y,
    k = 5,
    seed = 20260818
) {
  
  set.seed(seed)
  
  foldid <- integer(
    length(y)
  )
  
  for (class_now in c(0, 1)) {
    
    idx <- which(
      y == class_now
    )
    
    idx <- sample(
      idx,
      length(idx),
      replace = FALSE
    )
    
    fold_seq <- rep(
      seq_len(k),
      length.out = length(idx)
    )
    
    foldid[idx] <- fold_seq
  }
  
  if (any(foldid == 0)) {
    stop("Stratified fold allocation failed.")
  }
  
  fold_table <- data.frame(
    Fold = seq_len(k),
    N = vapply(
      seq_len(k),
      function(f) {
        sum(
          foldid == f
        )
      },
      numeric(1)
    ),
    Events = vapply(
      seq_len(k),
      function(f) {
        sum(
          y[
            foldid == f
          ] == 1
        )
      },
      numeric(1)
    ),
    Non_events = vapply(
      seq_len(k),
      function(f) {
        sum(
          y[
            foldid == f
          ] == 0
        )
      },
      numeric(1)
    ),
    stringsAsFactors = FALSE
  )
  
  if (any(fold_table$Events == 0)) {
    stop("At least one CV fold contains zero POUR events.")
  }
  
  list(
    foldid = foldid,
    fold_table = fold_table
  )
}


# ----------------------------------------------------------
# 16.4F Extract out-of-fold performance at lambda.min / 1se
#
# cv.glmnet(keep=TRUE) stores PREVALIDATED LINEAR PREDICTORS
# (logit scale for binomial models) in fit.preval.
#
# assess.glmnet() is retained for AUC and binomial deviance.
#
# IMPORTANT BRIER-SCORE DEFINITION:
# glmnet's binomial "mse" sums squared errors over BOTH
# class probabilities. For binary 0/1 outcomes this equals
# 2 * mean((y - p)^2).
#
# To keep the Brier score on the same conventional binary
# scale used in Sections 12-13, this section calculates:
#
#     Brier = mean((y - p)^2)
#
# where p = plogis(fit.preval).
#
# A numerical cross-check against glmnet$mse / 2 is included.
# ----------------------------------------------------------

extract_cv_performance_16 <- function(
    cvfit,
    y,
    rule
) {
  
  lambda_value <- if (
    rule == "lambda.min"
  ) {
    cvfit$lambda.min
  } else {
    cvfit$lambda.1se
  }
  
  lambda_index <- which.min(
    abs(
      cvfit$lambda -
        lambda_value
    )
  )
  
  pred_prevalidated <- cvfit$fit.preval[
    ,
    lambda_index
  ]
  
  if (
    anyNA(
      pred_prevalidated
    ) ||
    any(
      !is.finite(
        pred_prevalidated
      )
    )
  ) {
    return(
      data.frame(
        Lambda_rule = rule,
        Lambda = lambda_value,
        CV_AUC = NA_real_,
        CV_Brier = NA_real_,
        CV_Deviance = NA_real_,
        stringsAsFactors = FALSE
      )
    )
  }
  
  # fit.preval is on the link/logit scale for binomial glmnet.
  pred_probability <- stats::plogis(
    pred_prevalidated
  )
  
  if (
    anyNA(
      pred_probability
    ) ||
    any(
      !is.finite(
        pred_probability
      )
    ) ||
    any(
      pred_probability < 0 |
      pred_probability > 1
    )
  ) {
    return(
      data.frame(
        Lambda_rule = rule,
        Lambda = lambda_value,
        CV_AUC = NA_real_,
        CV_Brier = NA_real_,
        CV_Deviance = NA_real_,
        stringsAsFactors = FALSE
      )
    )
  }
  
  assess_now <- tryCatch(
    
    glmnet::assess.glmnet(
      matrix(
        pred_prevalidated,
        ncol = 1
      ),
      newy = y,
      family = "binomial"
    ),
    
    error = function(e) NULL
  )
  
  if (is.null(assess_now)) {
    
    return(
      data.frame(
        Lambda_rule = rule,
        Lambda = lambda_value,
        CV_AUC = NA_real_,
        CV_Brier = NA_real_,
        CV_Deviance = NA_real_,
        stringsAsFactors = FALSE
      )
    )
  }
  
  # Conventional binary Brier score, same scale as Sections 12-13.
  brier_conventional <- mean(
    (
      as.numeric(y) -
        pred_probability
    )^2,
    na.rm = TRUE
  )
  
  # glmnet's binomial mse uses the sum over the two classes,
  # which should equal 2 * conventional binary Brier.
  brier_from_glmnet_mse <- as.numeric(
    assess_now$mse[1]
  ) / 2
  
  brier_check_difference <- abs(
    brier_conventional -
      brier_from_glmnet_mse
  )
  
  if (
    is.finite(
      brier_check_difference
    ) &&
    brier_check_difference > 1e-10
  ) {
    warning(
      paste0(
        "Conventional Brier score differs from glmnet mse/2 by ",
        signif(
          brier_check_difference,
          4
        ),
        ". Please check fit.preval scale or outcome coding."
      )
    )
  }
  
  data.frame(
    Lambda_rule = rule,
    Lambda = lambda_value,
    CV_AUC = as.numeric(
      assess_now$auc[1]
    ),
    CV_Brier = brier_conventional,
    CV_Deviance = as.numeric(
      assess_now$deviance[1]
    ),
    stringsAsFactors = FALSE
  )
}


# ----------------------------------------------------------
# 16.4G Extract coefficients at lambda.min / lambda.1se
# ----------------------------------------------------------

extract_penalized_coefficients_16 <- function(
    cvfit,
    rule,
    term_map,
    timepoint,
    method_name,
    alpha_value,
    imputation_id,
    primary_vars
) {
  
  coef_matrix <- as.matrix(
    stats::coef(
      cvfit,
      s = rule
    )
  )
  
  coef_table <- data.frame(
    Term = rownames(coef_matrix),
    Beta = as.numeric(
      coef_matrix[, 1]
    ),
    stringsAsFactors = FALSE
  )
  
  coef_table$Variable <- ifelse(
    coef_table$Term == "(Intercept)",
    "(Intercept)",
    NA_character_
  )
  
  for (j in seq_len(nrow(term_map))) {
    coef_table$Variable[
      coef_table$Term ==
        term_map$Term[j]
    ] <- term_map$Variable[j]
  }
  
  if (
    any(
      is.na(
        coef_table$Variable
      )
    )
  ) {
    stop(
      paste0(
        "Unable to map one or more penalized coefficients ",
        "back to the original variables."
      )
    )
  }
  
  coef_table$Timepoint <- timepoint
  coef_table$Method <- method_name
  coef_table$Alpha <- alpha_value
  coef_table$Imputation <- imputation_id
  coef_table$Lambda_rule <- rule
  coef_table$OR <- exp(
    coef_table$Beta
  )
  
  coef_table$Selected <- ifelse(
    coef_table$Term == "(Intercept)",
    NA,
    abs(
      coef_table$Beta
    ) > 1e-10
  )
  
  coef_table$Primary_Logistic_Predictor <- ifelse(
    coef_table$Variable %in%
      primary_vars,
    "Yes",
    "No"
  )
  
  coef_table[
    ,
    c(
      "Timepoint",
      "Method",
      "Alpha",
      "Imputation",
      "Lambda_rule",
      "Variable",
      "Term",
      "Beta",
      "OR",
      "Selected",
      "Primary_Logistic_Predictor"
    )
  ]
}


# ==========================================================
# 16.5 Validate row order/outcomes and define common folds
# ==========================================================

first_preop_16 <- mice::complete(
  imp_preoperative,
  action = 1
)

first_precat_16 <- mice::complete(
  imp_precatheter,
  action = 1
)


y_reference_16 <- coerce_binary_outcome_16(
  first_preop_16[[outcome_var]],
  variable_name = outcome_var
)


y_precat_reference_16 <- coerce_binary_outcome_16(
  first_precat_16[[outcome_var]],
  variable_name = outcome_var
)


if (
  !identical(
    y_reference_16,
    y_precat_reference_16
  )
) {
  stop(
    "The outcome vector differs between the preoperative and pre-catheter MICE objects."
  )
}


N_16 <- length(
  y_reference_16
)

Events_16 <- sum(
  y_reference_16 == 1
)


for (i in seq_len(20)) {
  
  data_pre_i <- mice::complete(
    imp_preoperative,
    action = i
  )
  
  data_precat_i <- mice::complete(
    imp_precatheter,
    action = i
  )
  
  y_pre_i <- coerce_binary_outcome_16(
    data_pre_i[[outcome_var]],
    variable_name = outcome_var
  )
  
  y_precat_i <- coerce_binary_outcome_16(
    data_precat_i[[outcome_var]],
    variable_name = outcome_var
  )
  
  if (
    !identical(
      y_pre_i,
      y_reference_16
    )
  ) {
    stop(
      paste0(
        "Outcome vector differs in preoperative imputation ",
        i,
        "."
      )
    )
  }
  
  if (
    !identical(
      y_precat_i,
      y_reference_16
    )
  ) {
    stop(
      paste0(
        "Outcome vector differs in pre-catheter imputation ",
        i,
        "."
      )
    )
  }
}


cv_k_16 <- 5
cv_seed_16 <- 20260818


fold_object_16 <- make_stratified_foldid_16(
  y = y_reference_16,
  k = cv_k_16,
  seed = cv_seed_16
)

foldid_16 <- fold_object_16$foldid
fold_distribution_16 <- fold_object_16$fold_table


cat(
  "\nSection 16 common CV fold distribution:\n"
)

print(
  fold_distribution_16
)


# ==========================================================
# 16.6 Define methods
# ==========================================================

method_table_16 <- data.frame(
  Method = c(
    "Ridge",
    "Elastic_Net",
    "LASSO"
  ),
  Alpha = c(
    0,
    0.5,
    1
  ),
  stringsAsFactors = FALSE
)


cat(
  "\nPenalized methods:\n"
)

print(
  method_table_16
)


# ==========================================================
# 16.7 Prepare design-matrix reference structures
# ==========================================================

x_preop_reference_16 <- build_design_matrix_16(
  data_now = first_preop_16,
  predictor_vars = preoperative_full_vars_16
)

x_precat_reference_16 <- build_design_matrix_16(
  data_now = first_precat_16,
  predictor_vars = precatheter_full_vars_16
)


preop_design_columns_16 <- colnames(
  x_preop_reference_16
)

precat_design_columns_16 <- colnames(
  x_precat_reference_16
)


preop_term_map_16 <- make_term_map_16(
  data_now = first_preop_16,
  predictor_vars = preoperative_full_vars_16
)

precat_term_map_16 <- make_term_map_16(
  data_now = first_precat_16,
  predictor_vars = precatheter_full_vars_16
)


# ==========================================================
# 16.8 Core fitting function for one time point
# ==========================================================

fit_penalized_timepoint_16 <- function(
    imp_object,
    predictor_vars,
    expected_columns,
    term_map,
    timepoint_name,
    primary_vars
) {
  
  performance_list <- list()
  coefficient_list <- list()
  fit_list <- list()
  qc_list <- list()
  
  result_counter <- 1L
  
  
  for (i in seq_len(imp_object$m)) {
    
    cat(
      "\n",
      timepoint_name,
      " - imputation ",
      i,
      "/",
      imp_object$m,
      "\n",
      sep = ""
    )
    
    flush.console()
    
    
    data_i <- mice::complete(
      imp_object,
      action = i
    )
    
    y_i <- coerce_binary_outcome_16(
      data_i[[outcome_var]],
      variable_name = outcome_var
    )
    
    x_i <- build_design_matrix_16(
      data_now = data_i,
      predictor_vars = predictor_vars,
      expected_columns = expected_columns
    )
    
    
    for (method_id in seq_len(nrow(method_table_16))) {
      
      method_now <- method_table_16$Method[
        method_id
      ]
      
      alpha_now <- method_table_16$Alpha[
        method_id
      ]
      
      
      cat(
        "  ",
        method_now,
        " (alpha=",
        alpha_now,
        ")\n",
        sep = ""
      )
      
      flush.console()
      
      
      captured_warnings <- character()
      
      
      cvfit_now <- tryCatch(
        
        withCallingHandlers(
          
          glmnet::cv.glmnet(
            x = x_i,
            y = y_i,
            family = "binomial",
            alpha = alpha_now,
            foldid = foldid_16,
            nfolds = cv_k_16,
            type.measure = "deviance",
            standardize = TRUE,
            intercept = TRUE,
            keep = TRUE,
            nlambda = 100,
            parallel = FALSE
          ),
          
          warning = function(w) {
            
            captured_warnings <<- c(
              captured_warnings,
              conditionMessage(w)
            )
            
            invokeRestart(
              "muffleWarning"
            )
          }
        ),
        
        error = function(e) e
      )
      
      
      if (
        inherits(
          cvfit_now,
          "error"
        )
      ) {
        
        qc_list[[result_counter]] <- data.frame(
          Timepoint = timepoint_name,
          Imputation = i,
          Method = method_now,
          Alpha = alpha_now,
          Success = FALSE,
          Warning_n = length(captured_warnings),
          Message = conditionMessage(cvfit_now),
          stringsAsFactors = FALSE
        )
        
        result_counter <- result_counter + 1L
        
        next
      }
      
      
      # Save the cv.glmnet object
      fit_key <- paste(
        timepoint_name,
        paste0("MI", i),
        method_now,
        sep = "__"
      )
      
      fit_list[[fit_key]] <- cvfit_now
      
      
      # ----------------------------------------------------
      # Performance at lambda.min and lambda.1se
      # ----------------------------------------------------
      
      perf_min <- extract_cv_performance_16(
        cvfit = cvfit_now,
        y = y_i,
        rule = "lambda.min"
      )
      
      perf_1se <- extract_cv_performance_16(
        cvfit = cvfit_now,
        y = y_i,
        rule = "lambda.1se"
      )
      
      perf_now <- dplyr::bind_rows(
        perf_min,
        perf_1se
      )
      
      perf_now$Timepoint <- timepoint_name
      perf_now$Imputation <- i
      perf_now$Method <- method_now
      perf_now$Alpha <- alpha_now
      
      
      # Number of nonzero coefficients
      for (r in seq_len(nrow(perf_now))) {
        
        rule_now <- perf_now$Lambda_rule[r]
        
        coef_now <- as.matrix(
          stats::coef(
            cvfit_now,
            s = rule_now
          )
        )
        
        perf_now$Nonzero_predictors[r] <- sum(
          abs(
            coef_now[
              rownames(coef_now) != "(Intercept)",
              1
            ]
          ) > 1e-10
        )
      }
      
      
      performance_list[[length(performance_list) + 1L]] <- perf_now[
        ,
        c(
          "Timepoint",
          "Imputation",
          "Method",
          "Alpha",
          "Lambda_rule",
          "Lambda",
          "Nonzero_predictors",
          "CV_AUC",
          "CV_Brier",
          "CV_Deviance"
        )
      ]
      
      
      # ----------------------------------------------------
      # Coefficients at lambda.min and lambda.1se
      # ----------------------------------------------------
      
      coef_min <- extract_penalized_coefficients_16(
        cvfit = cvfit_now,
        rule = "lambda.min",
        term_map = term_map,
        timepoint = timepoint_name,
        method_name = method_now,
        alpha_value = alpha_now,
        imputation_id = i,
        primary_vars = primary_vars
      )
      
      coef_1se <- extract_penalized_coefficients_16(
        cvfit = cvfit_now,
        rule = "lambda.1se",
        term_map = term_map,
        timepoint = timepoint_name,
        method_name = method_now,
        alpha_value = alpha_now,
        imputation_id = i,
        primary_vars = primary_vars
      )
      
      coefficient_list[[length(coefficient_list) + 1L]] <-
        dplyr::bind_rows(
          coef_min,
          coef_1se
        )
      
      
      qc_list[[result_counter]] <- data.frame(
        Timepoint = timepoint_name,
        Imputation = i,
        Method = method_now,
        Alpha = alpha_now,
        Success = TRUE,
        Warning_n = length(captured_warnings),
        Message = if (
          length(captured_warnings) == 0
        ) {
          ""
        } else {
          paste(
            unique(captured_warnings),
            collapse = " | "
          )
        },
        stringsAsFactors = FALSE
      )
      
      result_counter <- result_counter + 1L
    }
  }
  
  
  list(
    performance = dplyr::bind_rows(
      performance_list
    ),
    coefficients = dplyr::bind_rows(
      coefficient_list
    ),
    fits = fit_list,
    qc = dplyr::bind_rows(
      qc_list
    )
  )
}


# ==========================================================
# 16.9 Run both time points
# ==========================================================

cat(
  "\n>>> Fitting preoperative penalized models <<<\n"
)


preoperative_penalized_16 <- fit_penalized_timepoint_16(
  imp_object = imp_preoperative,
  predictor_vars = preoperative_full_vars_16,
  expected_columns = preop_design_columns_16,
  term_map = preop_term_map_16,
  timepoint_name = "Preoperative",
  primary_vars = primary_preop_vars_16
)


cat(
  "\n>>> Fitting pre-catheter-removal penalized models <<<\n"
)


precatheter_penalized_16 <- fit_penalized_timepoint_16(
  imp_object = imp_precatheter,
  predictor_vars = precatheter_full_vars_16,
  expected_columns = precat_design_columns_16,
  term_map = precat_term_map_16,
  timepoint_name = "Pre-catheter-removal",
  primary_vars = primary_precat_vars_16
)


# ==========================================================
# 16.10 Combine detailed results
# ==========================================================

performance_detail_16 <- dplyr::bind_rows(
  preoperative_penalized_16$performance,
  precatheter_penalized_16$performance
)

coefficient_detail_16 <- dplyr::bind_rows(
  preoperative_penalized_16$coefficients,
  precatheter_penalized_16$coefficients
)

qc_detail_16 <- dplyr::bind_rows(
  preoperative_penalized_16$qc,
  precatheter_penalized_16$qc
)


# ==========================================================
# 16.11 Performance summary across 20 imputations
# ==========================================================

performance_summary_16 <-
  performance_detail_16 |>
  dplyr::group_by(
    Timepoint,
    Method,
    Alpha,
    Lambda_rule
  ) |>
  dplyr::summarise(
    Valid_imputations = sum(
      is.finite(CV_AUC)
    ),
    
    Mean_lambda = safe_mean_16(
      Lambda
    ),
    
    SD_lambda = safe_sd_16(
      Lambda
    ),
    
    Mean_nonzero_predictors = safe_mean_16(
      Nonzero_predictors
    ),
    
    Mean_CV_AUC = safe_mean_16(
      CV_AUC
    ),
    
    SD_CV_AUC = safe_sd_16(
      CV_AUC
    ),
    
    Median_CV_AUC = safe_median_16(
      CV_AUC
    ),
    
    Mean_CV_Brier = safe_mean_16(
      CV_Brier
    ),
    
    SD_CV_Brier = safe_sd_16(
      CV_Brier
    ),
    
    Mean_CV_Deviance = safe_mean_16(
      CV_Deviance
    ),
    
    SD_CV_Deviance = safe_sd_16(
      CV_Deviance
    ),
    
    .groups = "drop"
  )


# ==========================================================
# 16.12 Coefficient / selection stability across imputations
# ==========================================================

coefficient_nonintercept_16 <- coefficient_detail_16[
  coefficient_detail_16$Variable != "(Intercept)",
  ,
  drop = FALSE
]


coefficient_summary_16 <-
  coefficient_nonintercept_16 |>
  dplyr::group_by(
    Timepoint,
    Method,
    Alpha,
    Lambda_rule,
    Variable,
    Primary_Logistic_Predictor
  ) |>
  dplyr::summarise(
    Evaluated_imputations = dplyr::n(),
    
    Selected_n = sum(
      Selected,
      na.rm = TRUE
    ),
    
    Selection_frequency = ifelse(
      unique(Method) == "Ridge",
      NA_real_,
      mean(
        Selected,
        na.rm = TRUE
      )
    ),
    
    Mean_Beta = safe_mean_16(
      Beta
    ),
    
    SD_Beta = safe_sd_16(
      Beta
    ),
    
    Median_Beta = safe_median_16(
      Beta
    ),
    
    Mean_OR = safe_mean_16(
      OR
    ),
    
    Positive_n = sum(
      Beta > 1e-10,
      na.rm = TRUE
    ),
    
    Negative_n = sum(
      Beta < -1e-10,
      na.rm = TRUE
    ),
    
    Direction_consistency = {
      nonzero_n <- sum(
        abs(Beta) > 1e-10,
        na.rm = TRUE
      )
      
      if (nonzero_n == 0) {
        NA_real_
      } else {
        max(
          sum(Beta > 1e-10, na.rm = TRUE),
          sum(Beta < -1e-10, na.rm = TRUE)
        ) / nonzero_n
      }
    },
    
    .groups = "drop"
  )


# ----------------------------------------------------------
# 16.12A LASSO / Elastic-Net selection-frequency table
# ----------------------------------------------------------

selection_frequency_16 <-
  coefficient_summary_16 |>
  dplyr::filter(
    Method %in%
      c(
        "LASSO",
        "Elastic_Net"
      )
  ) |>
  dplyr::arrange(
    Timepoint,
    Lambda_rule,
    dplyr::desc(
      Selection_frequency
    )
  )


# ==========================================================
# 16.13 Simple stability flags
#
# These thresholds are descriptive only:
# - >= 0.80 : highly stable
# - 0.50-0.79 : moderately stable
# - < 0.50 : unstable/occasional
# ==========================================================

selection_frequency_16$Stability_label <- ifelse(
  selection_frequency_16$Selection_frequency >= 0.80,
  "High (>=80%)",
  ifelse(
    selection_frequency_16$Selection_frequency >= 0.50,
    "Moderate (50-79%)",
    "Low (<50%)"
  )
)


# ==========================================================
# 16.14 QC summary
# ==========================================================

qc_summary_16 <-
  qc_detail_16 |>
  dplyr::group_by(
    Timepoint,
    Method,
    Alpha
  ) |>
  dplyr::summarise(
    Expected_fits = 20,
    Successful_fits = sum(
      Success
    ),
    Failed_fits = sum(
      !Success
    ),
    Total_warnings = sum(
      Warning_n
    ),
    .groups = "drop"
  )


if (
  any(
    qc_summary_16$Failed_fits > 0
  )
) {
  warning(
    "At least one penalized-regression fit failed. Check the QC sheet."
  )
}


# ==========================================================
# 16.15 Package versions / reproducibility
# ==========================================================

reproducibility_16 <- data.frame(
  Item = c(
    "R_version",
    "glmnet_version",
    "mice_version",
    "CV_folds",
    "CV_seed",
    "ElasticNet_alpha",
    "Lambda_selection_measure",
    "Primary_lambda_rules_reported"
  ),
  Value = c(
    R.version.string,
    as.character(
      utils::packageVersion(
        "glmnet"
      )
    ),
    as.character(
      utils::packageVersion(
        "mice"
      )
    ),
    as.character(
      cv_k_16
    ),
    as.character(
      cv_seed_16
    ),
    "0.5",
    "Binomial deviance",
    "lambda.min and lambda.1se"
  ),
  stringsAsFactors = FALSE
)


# ==========================================================
# 16.16 Save one consolidated Excel workbook
# ==========================================================

output_excel_16 <- file.path(
  table_dir,
  "Penalized_regression_sensitivity_analysis.xlsx"
)


openxlsx::write.xlsx(
  list(
    Candidate_vars =
      candidate_table_16,
    
    Methods =
      method_table_16,
    
    Fold_distribution =
      fold_distribution_16,
    
    Performance_summary =
      performance_summary_16,
    
    Performance_by_MI =
      performance_detail_16,
    
    Selection_frequency =
      selection_frequency_16,
    
    Coefficient_summary =
      coefficient_summary_16,
    
    Coefficients_by_MI =
      coefficient_detail_16,
    
    QC =
      qc_summary_16,
    
    Reproducibility =
      reproducibility_16
  ),
  
  output_excel_16,
  
  overwrite = TRUE
)


# ==========================================================
# 16.17 Save model objects
# ==========================================================

output_rds_16 <- file.path(
  model_dir,
  "Penalized_regression_sensitivity_objects.rds"
)


saveRDS(
  list(
    Preoperative =
      preoperative_penalized_16,
    
    Precatheter =
      precatheter_penalized_16,
    
    Candidate_variables =
      list(
        Preoperative =
          preoperative_full_vars_16,
        Precatheter =
          precatheter_full_vars_16
      ),
    
    Method_table =
      method_table_16,
    
    Foldid =
      foldid_16,
    
    Fold_distribution =
      fold_distribution_16,
    
    Performance_summary =
      performance_summary_16,
    
    Selection_frequency =
      selection_frequency_16,
    
    Coefficient_summary =
      coefficient_summary_16,
    
    Reproducibility =
      reproducibility_16
  ),
  
  output_rds_16
)


# ==========================================================
# 16.18 Save QC log
# ==========================================================

output_qc_16 <- file.path(
  log_dir,
  "Penalized_regression_sensitivity_QC.xlsx"
)


openxlsx::write.xlsx(
  list(
    QC_summary =
      qc_summary_16,
    
    QC_detail =
      qc_detail_16,
    
    Fold_distribution =
      fold_distribution_16,
    
    Reproducibility =
      reproducibility_16
  ),
  
  output_qc_16,
  
  overwrite = TRUE
)


# ==========================================================
# 16.19 Final output-file QC
# ==========================================================

required_files_16 <- c(
  output_excel_16,
  output_rds_16,
  output_qc_16
)


missing_files_16 <- required_files_16[
  !file.exists(
    required_files_16
  )
]


if (
  length(
    missing_files_16
  ) > 0
) {
  stop(
    paste0(
      "Section 16 failed to create: ",
      paste(
        basename(
          missing_files_16
        ),
        collapse = ", "
      )
    )
  )
}


# ==========================================================
# 16.20 Console output
# ==========================================================

cat(
  "\n====================================================\n"
)

cat(
  "Penalized-regression CV performance summary:\n"
)

print(
  performance_summary_16
)


cat(
  "\n====================================================\n"
)

cat(
  "LASSO / Elastic-Net selection frequencies:\n"
)

print(
  selection_frequency_16
)


cat(
  "\n====================================================\n"
)

cat(
  "Penalized-regression QC:\n"
)

print(
  qc_summary_16
)


cat(
  "\n>>> Section 16 completed <<<\n"
)


cat(
  "\nPlease inspect:\n",
  "1. 02_tables/Penalized_regression_sensitivity_analysis.xlsx\n",
  "2. 05_logs/Penalized_regression_sensitivity_QC.xlsx\n",
  "3. 03_models/Penalized_regression_sensitivity_objects.rds\n",
  sep = ""
)


############################################################
############################################################
############################################################
# 16A. Supplementary outputs for penalized-regression
#      sensitivity analysis
#
# Purpose:
# 1. DO NOT refit Ridge / Elastic Net / LASSO models.
# 2. Read the already locked Section 16 results.
# 3. Generate publication-ready supplementary outputs:
#      - Supplementary Table S4:
#        Cross-validated predictive performance
#      - Supplementary Table S5:
#        Predictor selection frequency and direction
#      - Supplementary Figure S1:
#        Selection-frequency heatmap
# 4. Figure output: TIFF only, 600 dpi, LZW compression.
############################################################

cat("\n>>> Section 16A started: Supplementary Tables S4-S5 + Figure S1 <<<\n")
flush.console()


# ==========================================================
# 16A.1 Required package / directory checks
# ==========================================================

if (!requireNamespace("openxlsx", quietly = TRUE)) {
  stop("Section 16A requires the openxlsx package.")
}

if (!exists("table_dir", inherits = TRUE)) {
  stop("Section 16A cannot find table_dir. Please run it in the same project environment.")
}

project_root_16A <- dirname(table_dir)

figure_dir_16A <- file.path(
  project_root_16A,
  "04_figures"
)

log_dir_16A <- if (exists("log_dir", inherits = TRUE)) {
  log_dir
} else {
  file.path(
    project_root_16A,
    "05_logs"
  )
}


dir.create(
  figure_dir_16A,
  recursive = TRUE,
  showWarnings = FALSE
)

dir.create(
  log_dir_16A,
  recursive = TRUE,
  showWarnings = FALSE
)


input_file_16A <- file.path(
  table_dir,
  "Penalized_regression_sensitivity_analysis.xlsx"
)

if (!file.exists(input_file_16A)) {
  stop(
    paste0(
      "Section 16A cannot find: ",
      input_file_16A,
      "\nPlease make sure the corrected Section 16 has been run successfully."
    )
  )
}

cat("16A.1 Input/output checks: passed.\n")


# ==========================================================
# 16A.2 Read locked Section 16 results
# ==========================================================

performance_16A <- openxlsx::read.xlsx(
  input_file_16A,
  sheet = "Performance_summary"
)

selection_16A <- openxlsx::read.xlsx(
  input_file_16A,
  sheet = "Selection_frequency"
)

candidate_16A <- openxlsx::read.xlsx(
  input_file_16A,
  sheet = "Candidate_vars"
)

reproducibility_16A <- openxlsx::read.xlsx(
  input_file_16A,
  sheet = "Reproducibility"
)


required_perf_cols_16A <- c(
  "Timepoint",
  "Method",
  "Alpha",
  "Lambda_rule",
  "Valid_imputations",
  "Mean_lambda",
  "SD_lambda",
  "Mean_nonzero_predictors",
  "Mean_CV_AUC",
  "SD_CV_AUC",
  "Mean_CV_Brier",
  "SD_CV_Brier",
  "Mean_CV_Deviance",
  "SD_CV_Deviance"
)

missing_perf_cols_16A <- setdiff(
  required_perf_cols_16A,
  names(performance_16A)
)

if (length(missing_perf_cols_16A) > 0) {
  stop(
    paste0(
      "Performance_summary is missing columns: ",
      paste(missing_perf_cols_16A, collapse = ", ")
    )
  )
}


required_sel_cols_16A <- c(
  "Timepoint",
  "Method",
  "Lambda_rule",
  "Variable",
  "Primary_Logistic_Predictor",
  "Evaluated_imputations",
  "Selected_n",
  "Selection_frequency",
  "Mean_Beta",
  "Positive_n",
  "Negative_n"
)

missing_sel_cols_16A <- setdiff(
  required_sel_cols_16A,
  names(selection_16A)
)

if (length(missing_sel_cols_16A) > 0) {
  stop(
    paste0(
      "Selection_frequency is missing columns: ",
      paste(missing_sel_cols_16A, collapse = ", ")
    )
  )
}


cat("16A.2 Locked Section 16 results loaded successfully.\n")


# ==========================================================
# 16A.3 Publication labels and ordering
# ==========================================================

variable_labels_16A <- c(
  bmi = "BMI",
  coronary_heart_disease = "Coronary heart disease",
  popq_bp = "POP-Q Bp",
  bowel_difficulty = "Bowel difficulty",
  incomplete_emptying = "Incomplete emptying",
  detrusor_thickening = "Detrusor thickening",
  valsalva_urethral_inclination_angle = "Valsalva urethral inclination angle",
  bladder_neck_descent = "Bladder neck descent",
  colpocleisis_type = "Colpocleisis type",
  concomitant_hysterectomy = "Concomitant hysterectomy",
  blood_loss_ml = "Estimated blood loss"
)


label_variable_16A <- function(x) {
  out <- unname(variable_labels_16A[x])
  out[is.na(out)] <- x[is.na(out)]
  out
}


time_order_16A <- c(
  "Preoperative",
  "Pre-catheter-removal"
)

method_order_16A <- c(
  "Ridge",
  "Elastic_Net",
  "LASSO"
)

lambda_order_16A <- c(
  "lambda.min",
  "lambda.1se"
)


method_display_16A <- c(
  Ridge = "Ridge",
  Elastic_Net = "Elastic Net",
  LASSO = "LASSO"
)

lambda_display_16A <- c(
  "lambda.min" = "lambda.min",
  "lambda.1se" = "lambda.1se"
)


# ==========================================================
# 16A.4 Supplementary Table S4
# Cross-validated predictive performance
# ==========================================================

performance_16A$Time_order <- match(
  performance_16A$Timepoint,
  time_order_16A
)

performance_16A$Method_order <- match(
  performance_16A$Method,
  method_order_16A
)

performance_16A$Lambda_order <- match(
  performance_16A$Lambda_rule,
  lambda_order_16A
)

performance_16A <- performance_16A[
  order(
    performance_16A$Time_order,
    performance_16A$Method_order,
    performance_16A$Lambda_order
  ),
]


candidate_n_16A <- table(
  candidate_16A$Timepoint
)

candidate_count_16A <- function(tp) {
  n <- as.numeric(
    candidate_n_16A[tp]
  )
  if (length(n) == 0 || is.na(n)) {
    NA_real_
  } else {
    n
  }
}


format_mean_sd_16A <- function(mean_value, sd_value, digits) {
  if (!is.finite(mean_value)) {
    return(NA_character_)
  }
  
  if (!is.finite(sd_value)) {
    return(
      formatC(
        mean_value,
        format = "f",
        digits = digits
      )
    )
  }
  
  paste0(
    formatC(
      mean_value,
      format = "f",
      digits = digits
    ),
    " (",
    formatC(
      sd_value,
      format = "f",
      digits = digits
    ),
    ")"
  )
}


S4_16A <- data.frame(
  `Time point` = ifelse(
    performance_16A$Timepoint == "Preoperative",
    "Preoperative",
    "Pre-catheter-removal"
  ),
  Method = unname(
    method_display_16A[performance_16A$Method]
  ),
  Alpha = performance_16A$Alpha,
  `Penalty rule` = unname(
    lambda_display_16A[performance_16A$Lambda_rule]
  ),
  `Valid imputations` = performance_16A$Valid_imputations,
  `Mean lambda (SD)` = mapply(
    format_mean_sd_16A,
    performance_16A$Mean_lambda,
    performance_16A$SD_lambda,
    MoreArgs = list(digits = 4),
    USE.NAMES = FALSE
  ),
  `Mean nonzero predictors` = round(
    performance_16A$Mean_nonzero_predictors,
    2
  ),
  `CV AUC, mean (SD)` = mapply(
    format_mean_sd_16A,
    performance_16A$Mean_CV_AUC,
    performance_16A$SD_CV_AUC,
    MoreArgs = list(digits = 3),
    USE.NAMES = FALSE
  ),
  `CV Brier score, mean (SD)` = mapply(
    format_mean_sd_16A,
    performance_16A$Mean_CV_Brier,
    performance_16A$SD_CV_Brier,
    MoreArgs = list(digits = 4),
    USE.NAMES = FALSE
  ),
  `CV deviance, mean (SD)` = mapply(
    format_mean_sd_16A,
    performance_16A$Mean_CV_Deviance,
    performance_16A$SD_CV_Deviance,
    MoreArgs = list(digits = 3),
    USE.NAMES = FALSE
  ),
  check.names = FALSE,
  stringsAsFactors = FALSE
)


# ==========================================================
# 16A.5 Supplementary Table S5
# Predictor selection frequency and coefficient direction
# ==========================================================

selection_cell_16A <- function(
    data,
    timepoint,
    variable,
    method,
    lambda_rule
) {
  
  row_now <- data[
    data$Timepoint == timepoint &
      data$Variable == variable &
      data$Method == method &
      data$Lambda_rule == lambda_rule,
    ,
    drop = FALSE
  ]
  
  if (nrow(row_now) == 0) {
    return(NA_character_)
  }
  
  paste0(
    sprintf(
      "%.0f%%",
      100 * row_now$Selection_frequency[1]
    ),
    " (",
    row_now$Selected_n[1],
    "/",
    row_now$Evaluated_imputations[1],
    ")"
  )
}


selection_value_16A <- function(
    data,
    timepoint,
    variable,
    method,
    lambda_rule
) {
  
  row_now <- data[
    data$Timepoint == timepoint &
      data$Variable == variable &
      data$Method == method &
      data$Lambda_rule == lambda_rule,
    ,
    drop = FALSE
  ]
  
  if (nrow(row_now) == 0) {
    return(NA_real_)
  }
  
  as.numeric(
    row_now$Selection_frequency[1]
  )
}


direction_for_variable_16A <- function(
    data,
    timepoint,
    variable
) {
  
  rows_now <- data[
    data$Timepoint == timepoint &
      data$Variable == variable &
      data$Selected_n > 0,
    ,
    drop = FALSE
  ]
  
  if (nrow(rows_now) == 0) {
    return("Not selected")
  }
  
  total_positive <- sum(
    rows_now$Positive_n,
    na.rm = TRUE
  )
  
  total_negative <- sum(
    rows_now$Negative_n,
    na.rm = TRUE
  )
  
  if (total_positive > 0 && total_negative == 0) {
    return("Positive")
  }
  
  if (total_negative > 0 && total_positive == 0) {
    return("Negative")
  }
  
  if (total_positive > 0 && total_negative > 0) {
    return("Mixed")
  }
  
  "Not selected"
}


S2_list_16A <- list()
row_counter_16A <- 1

for (tp in time_order_16A) {
  
  variables_tp <- candidate_16A$Variable[
    candidate_16A$Timepoint == tp
  ]
  
  for (var_now in variables_tp) {
    
    primary_rows <- selection_16A[
      selection_16A$Timepoint == tp &
        selection_16A$Variable == var_now,
      ,
      drop = FALSE
    ]
    
    primary_flag <- if (
      nrow(primary_rows) > 0 &&
      any(primary_rows$Primary_Logistic_Predictor == "Yes")
    ) {
      "Yes"
    } else {
      "No"
    }
    
    S2_list_16A[[row_counter_16A]] <- data.frame(
      `Time point` = tp,
      Predictor = label_variable_16A(var_now),
      `Primary Logistic predictor` = primary_flag,
      `Elastic Net lambda.min` = selection_cell_16A(
        selection_16A,
        tp,
        var_now,
        "Elastic_Net",
        "lambda.min"
      ),
      `Elastic Net lambda.1se` = selection_cell_16A(
        selection_16A,
        tp,
        var_now,
        "Elastic_Net",
        "lambda.1se"
      ),
      `LASSO lambda.min` = selection_cell_16A(
        selection_16A,
        tp,
        var_now,
        "LASSO",
        "lambda.min"
      ),
      `LASSO lambda.1se` = selection_cell_16A(
        selection_16A,
        tp,
        var_now,
        "LASSO",
        "lambda.1se"
      ),
      `Coefficient direction` = direction_for_variable_16A(
        selection_16A,
        tp,
        var_now
      ),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    
    row_counter_16A <- row_counter_16A + 1
  }
}

S5_16A <- do.call(
  rbind,
  S2_list_16A
)


# ==========================================================
# 16A.6 Supplementary Figure S1 data matrix
# ==========================================================

heatmap_columns_16A <- c(
  "Elastic Net\nlambda.min",
  "Elastic Net\nlambda.1se",
  "LASSO\nlambda.min",
  "LASSO\nlambda.1se"
)


make_heatmap_matrix_16A <- function(timepoint_name) {
  
  variables_tp <- candidate_16A$Variable[
    candidate_16A$Timepoint == timepoint_name
  ]
  
  mat <- matrix(
    NA_real_,
    nrow = length(variables_tp),
    ncol = 4
  )
  
  colnames(mat) <- heatmap_columns_16A
  rownames(mat) <- label_variable_16A(
    variables_tp
  )
  
  primary_flag <- logical(
    length(variables_tp)
  )
  
  for (i in seq_along(variables_tp)) {
    
    var_now <- variables_tp[i]
    
    mat[i, 1] <- selection_value_16A(
      selection_16A,
      timepoint_name,
      var_now,
      "Elastic_Net",
      "lambda.min"
    )
    
    mat[i, 2] <- selection_value_16A(
      selection_16A,
      timepoint_name,
      var_now,
      "Elastic_Net",
      "lambda.1se"
    )
    
    mat[i, 3] <- selection_value_16A(
      selection_16A,
      timepoint_name,
      var_now,
      "LASSO",
      "lambda.min"
    )
    
    mat[i, 4] <- selection_value_16A(
      selection_16A,
      timepoint_name,
      var_now,
      "LASSO",
      "lambda.1se"
    )
    
    primary_rows <- selection_16A[
      selection_16A$Timepoint == timepoint_name &
        selection_16A$Variable == var_now,
      ,
      drop = FALSE
    ]
    
    primary_flag[i] <- (
      nrow(primary_rows) > 0 &&
        any(primary_rows$Primary_Logistic_Predictor == "Yes")
    )
  }
  
  list(
    matrix = mat,
    primary = primary_flag,
    variables = variables_tp
  )
}


heat_preop_16A <- make_heatmap_matrix_16A(
  "Preoperative"
)

heat_precat_16A <- make_heatmap_matrix_16A(
  "Pre-catheter-removal"
)


# ==========================================================
# 16A.7 Heatmap drawing function
# ==========================================================

heat_palette_16A <- grDevices::colorRampPalette(
  c(
    "#F7FBFF",
    "#C6DBEF",
    "#6BAED6",
    "#2171B5",
    "#08306B"
  )
)(101)


heat_color_16A <- function(value) {
  
  if (!is.finite(value)) {
    return("grey95")
  }
  
  index <- round(
    max(
      0,
      min(
        1,
        value
      )
    ) * 100
  ) + 1
  
  heat_palette_16A[index]
}


draw_heatmap_panel_16A <- function(
    heat_object,
    panel_letter,
    panel_title
) {
  
  mat <- heat_object$matrix
  primary <- heat_object$primary
  
  n_rows <- nrow(mat)
  n_cols <- ncol(mat)
  
  par(
    mar = c(5.8, 18.0, 3.2, 1.8),
    family = "sans",
    xpd = NA
  )
  
  plot(
    NA,
    xlim = c(0.5, n_cols + 0.5),
    ylim = c(0.5, n_rows + 0.5),
    xaxs = "i",
    yaxs = "i",
    axes = FALSE,
    xlab = "",
    ylab = "",
    main = ""
  )
  
  for (i in seq_len(n_rows)) {
    
    y_now <- n_rows - i + 1
    
    for (j in seq_len(n_cols)) {
      
      value_now <- mat[i, j]
      
      rect(
        j - 0.5,
        y_now - 0.5,
        j + 0.5,
        y_now + 0.5,
        col = heat_color_16A(value_now),
        border = "white",
        lwd = 1.2
      )
      
      label_now <- if (
        is.finite(value_now)
      ) {
        sprintf(
          "%.0f%%",
          100 * value_now
        )
      } else {
        "NA"
      }
      
      text_color_now <- if (
        is.finite(value_now) &&
        value_now >= 0.60
      ) {
        "white"
      } else {
        "black"
      }
      
      text(
        j,
        y_now,
        labels = label_now,
        cex = 0.88,
        font = 2,
        col = text_color_now
      )
    }
  }
  
  axis(
    side = 1,
    at = seq_len(n_cols),
    labels = colnames(mat),
    tick = FALSE,
    line = -0.25,
    cex.axis = 0.92
  )
  
  usr_now <- par("usr")
  label_x <- usr_now[1] - 0.07 * diff(usr_now[1:2])
  
  for (i in seq_len(n_rows)) {
    
    y_now <- n_rows - i + 1
    
    text(
      x = label_x,
      y = y_now,
      labels = rownames(mat)[i],
      adj = 1,
      cex = 0.88,
      font = if (primary[i]) 2 else 1
    )
  }
  
  title(
    main = paste0(
      panel_letter,
      ". ",
      panel_title
    ),
    cex.main = 1.12,
    font.main = 2,
    line = 1.0
  )
  
  box(
    col = "grey60",
    lwd = 0.8
  )
}


# ==========================================================
# 16A.8 Save Supplementary Figure S1 (TIFF only)
# ==========================================================

figure_S1_file_16A <- file.path(
  figure_dir_16A,
  "Supplementary_Figure_S1_penalized_selection_heatmap.tiff"
)


grDevices::tiff(
  filename = figure_S1_file_16A,
  width = 11.5,
  height = 12.0,
  units = "in",
  res = 600,
  compression = "lzw",
  bg = "white"
)

layout(
  matrix(
    c(
      1,
      2,
      3
    ),
    ncol = 1
  ),
  heights = c(
    8.0,
    10.5,
    1.4
  )
)


draw_heatmap_panel_16A(
  heat_preop_16A,
  panel_letter = "A",
  panel_title = "Preoperative candidate predictors"
)


draw_heatmap_panel_16A(
  heat_precat_16A,
  panel_letter = "B",
  panel_title = "Pre-catheter-removal candidate predictors"
)


# ----------------------------------------------------------
# Color scale / footnote panel
# ----------------------------------------------------------

par(
  mar = c(1.5, 18.0, 0.4, 1.8),
  family = "sans",
  xpd = NA
)

plot(
  NA,
  xlim = c(0, 1),
  ylim = c(0, 1),
  axes = FALSE,
  xlab = "",
  ylab = ""
)

legend_x0_16A <- 0.20
legend_x1_16A <- 0.80
legend_y0_16A <- 0.48
legend_y1_16A <- 0.68

for (k in 0:99) {
  
  x_left <- legend_x0_16A +
    (legend_x1_16A - legend_x0_16A) * k / 100
  
  x_right <- legend_x0_16A +
    (legend_x1_16A - legend_x0_16A) * (k + 1) / 100
  
  rect(
    x_left,
    legend_y0_16A,
    x_right,
    legend_y1_16A,
    col = heat_palette_16A[k + 1],
    border = NA
  )
}

rect(
  legend_x0_16A,
  legend_y0_16A,
  legend_x1_16A,
  legend_y1_16A,
  border = "grey45"
)

text(
  legend_x0_16A,
  legend_y0_16A - 0.12,
  "0%",
  cex = 0.80
)

text(
  (legend_x0_16A + legend_x1_16A) / 2,
  legend_y0_16A - 0.12,
  "50%",
  cex = 0.80
)

text(
  legend_x1_16A,
  legend_y0_16A - 0.12,
  "100%",
  cex = 0.80
)

text(
  0.5,
  0.88,
  "Selection frequency across 20 multiply imputed datasets",
  cex = 0.90,
  font = 2
)

text(
  0.5,
  0.14,
  "Bold predictor labels indicate variables retained in the primary Logistic model.",
  cex = 0.78
)


grDevices::dev.off()

cat("16A.8 Supplementary Figure S1 saved.\n")


# ==========================================================
# 16A.9 Save Supplementary Tables S4-S5 workbook
# ==========================================================

supplementary_excel_16A <- file.path(
  table_dir,
  "Supplementary_Tables_S4_S5_penalized_regression.xlsx"
)


notes_16A <- data.frame(
  Item = c(
    "Table S4 title",
    "Table S5 title",
    "Figure S1 title",
    "Internal-validation framework",
    "Elastic Net alpha",
    "Penalty rules",
    "Brier score",
    "Ridge note",
    "Selection frequency",
    "Primary-model marker"
  ),
  Description = c(
    "Cross-validated predictive performance of penalized regression models in sensitivity analyses.",
    "Predictor selection frequency and coefficient direction across 20 multiply imputed datasets in Elastic Net and LASSO sensitivity analyses.",
    "Predictor selection-frequency heatmap for penalized regression sensitivity analyses.",
    "A common stratified 5-fold cross-validation scheme was used across multiply imputed datasets and modelling approaches.",
    "0.5",
    "lambda.min and lambda.1se",
    "Conventional binary Brier score calculated on out-of-fold predicted probabilities; lower values indicate better overall probabilistic accuracy.",
    "Ridge regression retains all candidate predictors and therefore is included in Table S4 for predictive performance but not in Table S5 or Figure S1 for variable-selection frequency.",
    "Displayed as percentage and selected/evaluated imputations, e.g. 100% (20/20).",
    "Yes indicates that the predictor was retained in the primary traditional Logistic model."
  ),
  stringsAsFactors = FALSE
)


wb_16A <- openxlsx::createWorkbook()

openxlsx::addWorksheet(
  wb_16A,
  "Table_S4"
)

openxlsx::addWorksheet(
  wb_16A,
  "Table_S5"
)

openxlsx::addWorksheet(
  wb_16A,
  "Notes"
)


# ----------------------------------------------------------
# Styles
# ----------------------------------------------------------

header_style_16A <- openxlsx::createStyle(
  textDecoration = "bold",
  fgFill = "#D9EAF7",
  border = "Bottom",
  borderStyle = "thin",
  halign = "center",
  valign = "center",
  wrapText = TRUE
)

body_style_16A <- openxlsx::createStyle(
  valign = "center",
  wrapText = TRUE
)

primary_style_16A <- openxlsx::createStyle(
  textDecoration = "bold",
  valign = "center",
  wrapText = TRUE
)

section_border_16A <- openxlsx::createStyle(
  border = "Top",
  borderStyle = "medium"
)


# ----------------------------------------------------------
# Table S4
# ----------------------------------------------------------

openxlsx::writeData(
  wb_16A,
  "Table_S4",
  S4_16A,
  headerStyle = header_style_16A
)

openxlsx::addStyle(
  wb_16A,
  "Table_S4",
  body_style_16A,
  rows = 2:(nrow(S4_16A) + 1),
  cols = 1:ncol(S4_16A),
  gridExpand = TRUE,
  stack = TRUE
)

openxlsx::freezePane(
  wb_16A,
  "Table_S4",
  firstRow = TRUE
)

openxlsx::setColWidths(
  wb_16A,
  "Table_S4",
  cols = 1:ncol(S4_16A),
  widths = c(
    22,
    16,
    8,
    14,
    14,
    18,
    20,
    18,
    22,
    20
  )
)

precat_start_S4_16A <- which(
  S4_16A$`Time point` == "Pre-catheter-removal"
)[1] + 1

if (is.finite(precat_start_S4_16A)) {
  openxlsx::addStyle(
    wb_16A,
    "Table_S4",
    section_border_16A,
    rows = precat_start_S4_16A,
    cols = 1:ncol(S4_16A),
    gridExpand = TRUE,
    stack = TRUE
  )
}


# ----------------------------------------------------------
# Table S5
# ----------------------------------------------------------

openxlsx::writeData(
  wb_16A,
  "Table_S5",
  S5_16A,
  headerStyle = header_style_16A
)

openxlsx::addStyle(
  wb_16A,
  "Table_S5",
  body_style_16A,
  rows = 2:(nrow(S5_16A) + 1),
  cols = 1:ncol(S5_16A),
  gridExpand = TRUE,
  stack = TRUE
)

primary_rows_S5_16A <- which(
  S5_16A$`Primary Logistic predictor` == "Yes"
) + 1

if (length(primary_rows_S5_16A) > 0) {
  openxlsx::addStyle(
    wb_16A,
    "Table_S5",
    primary_style_16A,
    rows = primary_rows_S5_16A,
    cols = 2,
    gridExpand = TRUE,
    stack = TRUE
  )
}

openxlsx::freezePane(
  wb_16A,
  "Table_S5",
  firstRow = TRUE
)

openxlsx::setColWidths(
  wb_16A,
  "Table_S5",
  cols = 1:ncol(S5_16A),
  widths = c(
    22,
    34,
    20,
    20,
    20,
    18,
    18,
    20
  )
)

precat_start_S5_16A <- which(
  S5_16A$`Time point` == "Pre-catheter-removal"
)[1] + 1

if (is.finite(precat_start_S5_16A)) {
  openxlsx::addStyle(
    wb_16A,
    "Table_S5",
    section_border_16A,
    rows = precat_start_S5_16A,
    cols = 1:ncol(S5_16A),
    gridExpand = TRUE,
    stack = TRUE
  )
}


# ----------------------------------------------------------
# Notes
# ----------------------------------------------------------

openxlsx::writeData(
  wb_16A,
  "Notes",
  notes_16A,
  headerStyle = header_style_16A
)

openxlsx::addStyle(
  wb_16A,
  "Notes",
  body_style_16A,
  rows = 2:(nrow(notes_16A) + 1),
  cols = 1:2,
  gridExpand = TRUE,
  stack = TRUE
)

openxlsx::setColWidths(
  wb_16A,
  "Notes",
  cols = 1:2,
  widths = c(
    30,
    100
  )
)


openxlsx::saveWorkbook(
  wb_16A,
  supplementary_excel_16A,
  overwrite = TRUE
)

cat("16A.9 Supplementary Tables S4-S5 workbook saved.\n")


# ==========================================================
# 16A.10 QC
# ==========================================================

expected_S4_rows_16A <- 2 * 3 * 2
expected_S5_rows_16A <- nrow(candidate_16A)

if (nrow(S4_16A) != expected_S4_rows_16A) {
  warning(
    paste0(
      "Table S4 has ",
      nrow(S4_16A),
      " rows; expected ",
      expected_S4_rows_16A,
      "."
    )
  )
}

if (nrow(S5_16A) != expected_S5_rows_16A) {
  warning(
    paste0(
      "Table S5 has ",
      nrow(S5_16A),
      " rows; expected ",
      expected_S5_rows_16A,
      "."
    )
  )
}

if (
  any(
    selection_16A$Selection_frequency < 0 |
    selection_16A$Selection_frequency > 1,
    na.rm = TRUE
  )
) {
  stop("Selection frequencies outside the 0-1 range were detected.")
}

required_output_16A <- c(
  supplementary_excel_16A,
  figure_S1_file_16A
)

missing_output_16A <- required_output_16A[
  !file.exists(
    required_output_16A
  )
]

if (length(missing_output_16A) > 0) {
  stop(
    paste0(
      "Section 16A failed to generate: ",
      paste(
        basename(missing_output_16A),
        collapse = ", "
      )
    )
  )
}


cat("\nSupplementary Table S4 preview:\n")
print(S4_16A)

cat("\nSupplementary Table S5 preview:\n")
print(S5_16A)

cat(
  "\n>>> Section 16A completed successfully <<<\n",
  "Generated files:\n",
  "1. 02_tables/Supplementary_Tables_S4_S5_penalized_regression.xlsx\n",
  "2. 04_figures/Supplementary_Figure_S1_penalized_selection_heatmap.tiff\n",
  sep = ""
)

############################################################
# 17. Reproducibility metadata for public code release
############################################################

reproducibility_metadata_file <- file.path(
  log_dir,
  "sessionInfo.txt"
)

utils::capture.output(
  sessionInfo(),
  file = reproducibility_metadata_file
)

cat(
  "\nPublic reproducibility script completed.\n",
  "R/session metadata saved to: ",
  reproducibility_metadata_file,
  "\n",
  sep = ""
)

