# ============================================================
# Employee_Performance_Logistic_Regression.R
# 직원 성과 이항 로지스틱 회귀분석 + Test 예측
#
# 종속변수: 성과 (0=저성과, 1=고성과)
# 기준집단: 결혼여부 M / 성별 F / 학력 1
#
# 주요 기능
# 1. Train 데이터 로지스틱 회귀분석
# 2. OR, 95% CI, 다중공선성 및 모형 적합도
# 3. Train 예측 성능 평가
# 4. Test 데이터 고성과 예측확률 산출
# 5. Test 실제성과 수동 입력 및 자동 검증
# 6. Excel 결과 14개 시트 저장
#
# 성능 지표
# Accuracy / Precision / Specificity / F1 / AUC
# Recall은 F1 계산에만 내부적으로 사용
# ============================================================


# ============================================================
# STEP 1. 패키지 설치 및 불러오기
# ============================================================

packages <- c(
  "readxl",
  "openxlsx",
  "car",
  "ResourceSelection",
  "pscl",
  "pROC",
  "lmtest"
)

for (pkg in packages) {
  
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(
      pkg,
      repos = "https://cloud.r-project.org"
    )
  }
  
  suppressPackageStartupMessages(
    library(pkg, character.only = TRUE)
  )
}


# ============================================================
# STEP 2. 공통 설정 및 함수
# ============================================================

target <- "성과"

predictors <- c(
  "결혼여부",
  "근무시간",
  "근태사고율",
  "나이",
  "리더역량",
  "성별",
  "연봉",
  "연차",
  "조직적합도",
  "직무만족도",
  "학력"
)

numeric_vars <- c(
  "근무시간",
  "근태사고율",
  "나이",
  "리더역량",
  "연봉",
  "연차",
  "조직적합도",
  "직무만족도"
)

categorical_vars <- c(
  "결혼여부",
  "성별",
  "학력"
)

threshold <- 0.5


# 안전한 나눗셈
safe_div <- function(a, b) {
  
  if (is.na(b) || b == 0) {
    return(NA_real_)
  }
  
  a / b
}


# 백분율 표시
pct <- function(x) {
  
  if (length(x) == 0 || is.na(x)) {
    return("산출 불가")
  }
  
  sprintf("%.1f%%", x * 100)
}


# 숫자 표시
num_text <- function(x, digits = 3) {
  
  if (length(x) == 0 || is.na(x)) {
    return("산출 불가")
  }
  
  sprintf(paste0("%.", digits, "f"), x)
}


# AUC 해석
auc_grade <- function(x) {
  
  if (length(x) == 0 || is.na(x)) {
    return("평가 불가")
  }
  
  if (x >= 0.9) return("매우 우수")
  if (x >= 0.8) return("우수")
  if (x >= 0.7) return("보통")
  if (x >= 0.6) return("낮음")
  
  return("매우 낮음")
}


# ------------------------------------------------------------
# 파일 불러오기
# ------------------------------------------------------------

read_data <- function(path) {
  
  ext <- tolower(tools::file_ext(path))
  
  if (ext %in% c("xlsx", "xls")) {
    
    df <- as.data.frame(
      readxl::read_excel(path)
    )
    
  } else if (ext == "csv") {
    
    df <- tryCatch(
      read.csv(
        path,
        fileEncoding = "UTF-8-BOM",
        check.names = FALSE
      ),
      error = function(e) {
        read.csv(
          path,
          fileEncoding = "CP949",
          check.names = FALSE
        )
      }
    )
    
  } else {
    
    stop("Excel 또는 CSV 파일만 지원합니다.")
  }
  
  if (nrow(df) == 0) {
    stop("선택한 파일에 데이터가 없습니다.")
  }
  
  return(df)
}


# ------------------------------------------------------------
# 데이터 전처리
# ------------------------------------------------------------

preprocess_data <- function(
    df,
    is_train = TRUE,
    train_levels = NULL
) {
  
  missing_vars <- setdiff(
    predictors,
    names(df)
  )
  
  if (length(missing_vars) > 0) {
    
    stop(
      paste(
        "누락된 독립변수:",
        paste(missing_vars, collapse = ", ")
      )
    )
  }
  
  if (is_train && !("성과" %in% names(df))) {
    stop("Train 파일에는 성과 변수가 필요합니다.")
  }
  
  # Test 파일에도 실제 성과 필드 필요
  if (!is_train && !("성과" %in% names(df))) {
    stop("Test 파일에는 실제 성과 변수가 필요합니다.")
  }
  
  # 응답자 번호 생성
  df$`응답자 번호` <- seq_len(nrow(df))
  
  # 문자 공백 제거
  cols <- intersect(
    c(target, predictors),
    names(df)
  )
  
  for (v in cols) {
    
    if (is.character(df[[v]])) {
      
      df[[v]] <- trimws(df[[v]])
      
      df[[v]][df[[v]] == ""] <- NA
    }
  }
  
  # 성과 0/1 변환
  if ("성과" %in% names(df)) {
    
    raw_y <- as.character(df$성과)
    
    df$성과 <- suppressWarnings(
      as.numeric(raw_y)
    )
    
    if (any(
      !is.na(df$성과) &
      !df$성과 %in% c(0, 1)
    )) {
      
      stop("성과는 0 또는 1이어야 합니다.")
    }
    
    if (any(
      !is.na(raw_y) &
      is.na(df$성과)
    )) {
      
      stop(
        "성과 변수에 숫자로 변환할 수 없는 값이 있습니다."
      )
    }
  }
  
  # 연속형 변수
  for (v in numeric_vars) {
    
    raw_x <- gsub(
      ",",
      "",
      as.character(df[[v]])
    )
    
    df[[v]] <- suppressWarnings(
      as.numeric(raw_x)
    )
    
    if (any(
      !is.na(raw_x) &
      is.na(df[[v]])
    )) {
      
      warning(
        paste(
          v,
          "변수에 숫자로 변환할 수 없는 값이 있습니다."
        )
      )
    }
  }
  
  # 범주형 변수
  if (is_train) {
    
    df$결혼여부 <- factor(df$결혼여부)
    df$성별 <- factor(df$성별)
    
    edu_raw <- as.character(df$학력)
    
    if (any(
      !is.na(edu_raw) &
      !edu_raw %in% c("1", "2", "3")
    )) {
      
      stop("학력은 1, 2, 3으로 코딩해야 합니다.")
    }
    
    df$학력 <- factor(
      edu_raw,
      levels = c("1", "2", "3")
    )
    
    if (!("M" %in% levels(df$결혼여부))) {
      stop("결혼여부 기준집단 M이 없습니다.")
    }
    
    if (!("F" %in% levels(df$성별))) {
      stop("성별 기준집단 F가 없습니다.")
    }
    
    df$결혼여부 <- relevel(
      df$결혼여부,
      ref = "M"
    )
    
    df$성별 <- relevel(
      df$성별,
      ref = "F"
    )
    
  } else {
    
    for (v in categorical_vars) {
      
      raw_value <- as.character(df[[v]])
      
      unknown <- !is.na(raw_value) &
        !raw_value %in% train_levels[[v]]
      
      if (any(unknown)) {
        
        warning(
          paste(
            "Test 파일의",
            v,
            "변수에 Train에 없는 범주가 있습니다.",
            "해당 행은 예측에서 제외됩니다."
          )
        )
      }
      
      df[[v]] <- factor(
        raw_value,
        levels = train_levels[[v]]
      )
    }
  }
  
  return(df)
}


# ------------------------------------------------------------
# 예측 결과 필드 순서
# ------------------------------------------------------------

reorder_prediction <- function(df) {
  
  cols <- names(df)
  
  cols <- cols[cols != "응답자 번호"]
  
  pos <- match(
    "고성과_예측확률",
    cols
  )
  
  if (is.na(pos)) {
    stop("고성과_예측확률 열을 찾을 수 없습니다.")
  }
  
  cols <- append(
    cols,
    "응답자 번호",
    after = pos - 1
  )
  
  return(
    df[, cols, drop = FALSE]
  )
}


# ============================================================
# STEP 3. Train 파일 선택
# ============================================================

cat("\n[1/2] Train 파일을 선택하세요.\n")

train_path <- file.choose()

train_raw <- read_data(train_path)

output_dir <- dirname(train_path)

train_data <- preprocess_data(
  train_raw,
  is_train = TRUE
)

analysis_data <- train_data[
  complete.cases(
    train_data[
      ,
      c(target, predictors),
      drop = FALSE
    ]
  ),
  ,
  drop = FALSE
]

analysis_data <- droplevels(analysis_data)

if (nrow(analysis_data) == 0) {
  stop("Train 분석 가능 데이터가 없습니다.")
}

if (length(unique(analysis_data$성과)) != 2) {
  stop("Train 성과에 0과 1이 모두 있어야 합니다.")
}

if (
  nlevels(analysis_data$결혼여부) < 2 ||
  !("M" %in% levels(analysis_data$결혼여부))
) {
  
  stop(
    "결혼여부 M과 비교집단이 모두 필요합니다."
  )
}

if (
  nlevels(analysis_data$성별) < 2 ||
  !("F" %in% levels(analysis_data$성별))
) {
  
  stop(
    "성별 F와 비교집단이 모두 필요합니다."
  )
}

if (!all(
  c("1", "2", "3") %in%
  levels(analysis_data$학력)
)) {
  
  stop(
    "학력 1, 2, 3 집단이 모두 필요합니다."
  )
}

train_levels <- lapply(
  analysis_data[
    ,
    categorical_vars,
    drop = FALSE
  ],
  levels
)

cat(
  "\nTrain 전체:",
  nrow(train_data)
)

cat(
  "\nTrain 분석:",
  nrow(analysis_data),
  "\n"
)


# ============================================================
# STEP 4. 기술통계
# ============================================================

descriptive <- do.call(
  rbind,
  lapply(
    c("성과", numeric_vars),
    function(v) {
      
      x <- analysis_data[[v]]
      
      data.frame(
        변수 = v,
        N = length(x),
        평균 = mean(x),
        표준편차 = sd(x),
        최소값 = min(x),
        최대값 = max(x)
      )
    }
  )
)

categorical_summary <- do.call(
  rbind,
  lapply(
    categorical_vars,
    function(v) {
      
      freq <- as.data.frame(
        table(analysis_data[[v]])
      )
      
      names(freq) <- c(
        "범주",
        "빈도"
      )
      
      freq$변수 <- v
      
      freq$비율 <- round(
        freq$빈도 /
          sum(freq$빈도) * 100,
        2
      )
      
      freq[
        ,
        c(
          "변수",
          "범주",
          "빈도",
          "비율"
        )
      ]
    }
  )
)


# ============================================================
# STEP 5. 이항 로지스틱 회귀모형 학습
# ============================================================

formula_logit <- as.formula(
  paste(
    "성과 ~",
    paste(
      predictors,
      collapse = " + "
    )
  )
)

model <- glm(
  formula_logit,
  data = analysis_data,
  family = binomial(
    link = "logit"
  ),
  control = glm.control(
    maxit = 100
  )
)

if (!model$converged) {
  stop("로지스틱 회귀모형이 수렴하지 않았습니다.")
}

if (anyNA(coef(model))) {
  
  stop(
    paste(
      "회귀계수 일부를 추정할 수 없습니다.",
      "변수 중복 또는 완전 공선성을 확인하세요."
    )
  )
}

cat("\n===== 회귀모형 학습 완료 =====\n")

print(summary(model))


# ============================================================
# STEP 6. OR 및 95% 신뢰구간
# ============================================================

coef_table <- summary(model)$coefficients

beta <- coef(model)

se <- coef_table[, "Std. Error"]

or_results <- data.frame(
  변수 = names(beta),
  B = as.numeric(beta),
  SE = as.numeric(se),
  Wald_Z = as.numeric(
    coef_table[, "z value"]
  ),
  P_value = as.numeric(
    coef_table[, "Pr(>|z|)"]
  ),
  OR = exp(as.numeric(beta)),
  OR_95CI_Lower = exp(
    as.numeric(beta - 1.96 * se)
  ),
  OR_95CI_Upper = exp(
    as.numeric(beta + 1.96 * se)
  ),
  row.names = NULL
)

or_results$유의성 <- ifelse(
  is.na(or_results$P_value),
  "확인필요",
  ifelse(
    or_results$P_value < 0.001,
    "***",
    ifelse(
      or_results$P_value < 0.01,
      "**",
      ifelse(
        or_results$P_value < 0.05,
        "*",
        "ns"
      )
    )
  )
)

or_results$해석 <- ifelse(
  or_results$OR > 1,
  "고성과 오즈 증가",
  ifelse(
    or_results$OR < 1,
    "고성과 오즈 감소",
    "변화 없음"
  )
)


# ============================================================
# STEP 7. 다중공선성
# ============================================================

vif_raw <- car::vif(model)

if (is.matrix(vif_raw)) {
  
  vif_results <- data.frame(
    변수 = rownames(vif_raw),
    GVIF = vif_raw[, "GVIF"],
    Df = vif_raw[, "Df"],
    GVIF_Adjusted =
      vif_raw[, "GVIF^(1/(2*Df))"],
    row.names = NULL
  )
  
  vif_results$Equivalent_VIF <-
    vif_results$GVIF_Adjusted^2
  
} else {
  
  vif_results <- data.frame(
    변수 = names(vif_raw),
    GVIF = as.numeric(vif_raw),
    Df = 1,
    GVIF_Adjusted = sqrt(
      as.numeric(vif_raw)
    ),
    Equivalent_VIF =
      as.numeric(vif_raw),
    row.names = NULL
  )
}

vif_results$판정 <- ifelse(
  vif_results$Equivalent_VIF >= 10,
  "높음",
  ifelse(
    vif_results$Equivalent_VIF >= 5,
    "주의",
    "양호"
  )
)


# ============================================================
# STEP 8. 모형 적합도
# ============================================================

null_model <- glm(
  성과 ~ 1,
  data = analysis_data,
  family = binomial
)

lr_test <- lmtest::lrtest(
  null_model,
  model
)

lr_stat <- as.numeric(
  lr_test$Chisq[nrow(lr_test)]
)

lr_df <- as.numeric(
  lr_test$Df[nrow(lr_test)]
)

lr_p <- as.numeric(
  lr_test$`Pr(>Chisq)`[
    nrow(lr_test)
  ]
)

pseudo_r2 <- pscl::pR2(model)

n <- nrow(analysis_data)

ll_full <- as.numeric(
  logLik(model)
)

ll_null <- as.numeric(
  logLik(null_model)
)

cox_snell <- 1 - exp(
  (2 / n) *
    (ll_null - ll_full)
)

nagelkerke <- cox_snell /
  (1 - exp(2 * ll_null / n))

g_groups <- min(
  10,
  length(unique(fitted(model)))
)

hl <- if (g_groups >= 3) {
  
  tryCatch(
    ResourceSelection::hoslem.test(
      analysis_data$성과,
      fitted(model),
      g = g_groups
    ),
    error = function(e) NULL
  )
  
} else {
  
  NULL
}

hl_chisq <- if (!is.null(hl)) {
  as.numeric(hl$statistic)
} else {
  NA_real_
}

hl_df <- if (!is.null(hl)) {
  as.numeric(hl$parameter)
} else {
  NA_real_
}

hl_p <- if (!is.null(hl)) {
  as.numeric(hl$p.value)
} else {
  NA_real_
}

model_fit <- data.frame(
  지표 = c(
    "전체 Train 표본 수",
    "분석 Train 표본 수",
    "제외 표본 수",
    "고성과자 수",
    "저성과자 수",
    "AIC",
    "BIC",
    "Log Likelihood",
    "Null Deviance",
    "Residual Deviance",
    "LR Chi-square",
    "LR df",
    "LR p-value",
    "McFadden R2",
    "Cox-Snell R2",
    "Nagelkerke R2",
    "Hosmer-Lemeshow Chi-square",
    "Hosmer-Lemeshow df",
    "Hosmer-Lemeshow p-value"
  ),
  값 = c(
    nrow(train_data),
    n,
    nrow(train_data) - n,
    sum(analysis_data$성과 == 1),
    sum(analysis_data$성과 == 0),
    AIC(model),
    BIC(model),
    ll_full,
    model$null.deviance,
    model$deviance,
    lr_stat,
    lr_df,
    lr_p,
    as.numeric(
      pseudo_r2["McFadden"]
    ),
    cox_snell,
    nagelkerke,
    hl_chisq,
    hl_df,
    hl_p
  )
)


# ============================================================
# STEP 9. 공통 예측 함수
# ============================================================

predict_employees <- function(
    df,
    model,
    threshold
) {
  
  df$고성과_예측확률 <- NA_real_
  df$예측성과 <- NA_integer_
  df$예측분류 <- NA_character_
  df$예측정확여부 <- NA_character_
  
  valid_rows <- complete.cases(
    df[
      ,
      predictors,
      drop = FALSE
    ]
  )
  
  if (sum(valid_rows) > 0) {
    
    prob <- predict(
      model,
      newdata = df[
        valid_rows,
        ,
        drop = FALSE
      ],
      type = "response"
    )
    
    df$고성과_예측확률[
      valid_rows
    ] <- prob
    
    df$예측성과[
      valid_rows
    ] <- ifelse(
      prob >= threshold,
      1L,
      0L
    )
    
    df$예측분류[
      valid_rows
    ] <- ifelse(
      df$예측성과[
        valid_rows
      ] == 1,
      "고성과",
      "저성과"
    )
  }
  
  if ("성과" %in% names(df)) {
    
    eval_rows <- valid_rows &
      !is.na(df$성과) &
      !is.na(df$예측성과)
    
    df$예측정확여부[
      eval_rows
    ] <- ifelse(
      df$성과[eval_rows] ==
        df$예측성과[eval_rows],
      "정확",
      "오분류"
    )
  }
  
  return(
    reorder_prediction(df)
  )
}


# ============================================================
# STEP 10. 공통 예측 성능 평가 함수
# ============================================================

evaluate_model <- function(
    df,
    threshold
) {
  
  # Recall 제외, Sensitivity 유지
  metric_names <- c(
    "Accuracy",
    "Precision",
    "Sensitivity",
    "Specificity",
    "F1",
    "AUC"
  )
  
  metrics <- setNames(
    rep(
      NA_real_,
      length(metric_names)
    ),
    metric_names
  )
  
  has_actual <- "성과" %in% names(df)
  
  valid <- if (has_actual) {
    
    !is.na(df$성과) &
      !is.na(df$예측성과) &
      !is.na(df$고성과_예측확률)
    
  } else {
    
    rep(FALSE, nrow(df))
  }
  
  eval_n <- sum(valid)
  
  TP <- TN <- FP <- FN <- NA_integer_
  
  roc_obj <- NULL
  
  auc_ci <- c(
    NA_real_,
    NA_real_,
    NA_real_
  )
  
  if (eval_n > 0) {
    
    y <- df$성과[valid]
    
    yhat <- df$예측성과[valid]
    
    TP <- sum(
      y == 1 & yhat == 1
    )
    
    TN <- sum(
      y == 0 & yhat == 0
    )
    
    FP <- sum(
      y == 0 & yhat == 1
    )
    
    FN <- sum(
      y == 1 & yhat == 0
    )
    
    # Accuracy
    metrics["Accuracy"] <- safe_div(
      TP + TN,
      eval_n
    )
    
    # Precision
    metrics["Precision"] <- safe_div(
      TP,
      TP + FP
    )
    
    # Sensitivity
    metrics["Sensitivity"] <- safe_div(
      TP,
      TP + FN
    )
    
    # Specificity
    metrics["Specificity"] <- safe_div(
      TN,
      TN + FP
    )
    
    # F1 Score
    p <- metrics["Precision"]
    s <- metrics["Sensitivity"]
    
    if (!is.na(p) && !is.na(s)) {
      
      metrics["F1"] <- if (
        p + s == 0
      ) {
        
        0
        
      } else {
        
        2 * p * s / (p + s)
      }
    }
    
    # ROC-AUC
    if (length(unique(y)) == 2) {
      
      roc_obj <- pROC::roc(
        response = y,
        predictor =
          df$고성과_예측확률[valid],
        levels = c(0, 1),
        direction = "<",
        quiet = TRUE
      )
      
      metrics["AUC"] <- as.numeric(
        pROC::auc(roc_obj)
      )
      
      auc_ci <- tryCatch(
        as.numeric(
          pROC::ci.auc(roc_obj)
        ),
        error = function(e) {
          c(
            NA_real_,
            NA_real_,
            NA_real_
          )
        }
      )
    }
  }
  
  # 혼동행렬
  confusion <- data.frame(
    실제성과 = c(
      "고성과(1)",
      "저성과(0)"
    ),
    예측_고성과 = c(
      TP,
      FP
    ),
    예측_저성과 = c(
      FN,
      TN
    )
  )
  
  # 예측 성능 결과
  performance <- data.frame(
    지표 = c(
      "전체 응답자",
      "예측 가능 응답자",
      "실제 성과 비교 가능 응답자",
      "Threshold",
      "Accuracy",
      "Precision",
      "Sensitivity",
      "Specificity",
      "F1 Score",
      "ROC-AUC",
      "AUC 95% CI Lower",
      "AUC 95% CI Upper",
      "TP",
      "TN",
      "FP",
      "FN"
    ),
    결과 = c(
      nrow(df),
      sum(
        !is.na(df$고성과_예측확률)
      ),
      eval_n,
      threshold,
      metrics["Accuracy"],
      metrics["Precision"],
      metrics["Sensitivity"],
      metrics["Specificity"],
      metrics["F1"],
      metrics["AUC"],
      auc_ci[1],
      auc_ci[3],
      TP,
      TN,
      FP,
      FN
    ),
    row.names = NULL
  )
  
  performance$`결과 해석` <- c(
    "파일에 포함된 전체 응답자 수",
    "독립변수가 완전하여 예측 가능한 응답자 수",
    "실제 성과와 예측 성과를 비교할 수 있는 응답자 수",
    paste0(
      "고성과 분류 기준: ",
      pct(threshold)
    ),
    paste0(
      "전체 중 정확히 분류한 비율: ",
      pct(metrics["Accuracy"])
    ),
    paste0(
      "고성과 예측 중 실제 고성과 비율: ",
      pct(metrics["Precision"])
    ),
    paste0(
      "실제 고성과자 탐지율: ",
      pct(metrics["Sensitivity"])
    ),
    paste0(
      "실제 저성과자 탐지율: ",
      pct(metrics["Specificity"])
    ),
    paste0(
      "정밀도와 민감도의 조화평균: ",
      num_text(metrics["F1"])
    ),
    paste0(
      "고성과 구분 능력: ",
      auc_grade(metrics["AUC"]),
      " (",
      num_text(metrics["AUC"]),
      ")"
    ),
    paste0(
      "AUC 신뢰구간 하한: ",
      num_text(auc_ci[1])
    ),
    paste0(
      "AUC 신뢰구간 상한: ",
      num_text(auc_ci[3])
    ),
    paste0(
      "고성과를 정확히 예측한 인원: ",
      TP
    ),
    paste0(
      "저성과를 정확히 예측한 인원: ",
      TN
    ),
    paste0(
      "저성과를 고성과로 잘못 예측: ",
      FP
    ),
    paste0(
      "고성과를 저성과로 잘못 예측: ",
      FN
    )
  )
  
  return(
    list(
      metrics = metrics,
      performance = performance,
      confusion = confusion,
      roc = roc_obj
    )
  )
}


# ============================================================
# STEP 11. Train 예측 및 성능 평가
# ============================================================

train_prediction <- predict_employees(
  train_data,
  model,
  threshold
)

train_eval <- evaluate_model(
  train_prediction,
  threshold
)

roc_points <- if (
  !is.null(train_eval$roc)
) {
  
  data.frame(
    Threshold =
      train_eval$roc$thresholds,
    Sensitivity =
      train_eval$roc$sensitivities,
    Specificity =
      train_eval$roc$specificities,
    FPR =
      1 - train_eval$roc$specificities,
    TPR =
      train_eval$roc$sensitivities
  )
  
} else {
  
  data.frame(
    Threshold = numeric(),
    Sensitivity = numeric(),
    Specificity = numeric(),
    FPR = numeric(),
    TPR = numeric()
  )
}


# ============================================================
# STEP 12. Test 파일 선택
# ============================================================

cat("\n[2/2] Test 파일을 선택하세요.\n")

test_path <- file.choose()

if (
  normalizePath(train_path) ==
  normalizePath(test_path)
) {
  
  stop(
    "Train과 Test는 서로 다른 파일이어야 합니다."
  )
}

test_raw <- read_data(test_path)

test_data <- preprocess_data(
  test_raw,
  is_train = FALSE,
  train_levels = train_levels
)

cat(
  "\nTest 전체:",
  nrow(test_data),
  "\n"
)


# ============================================================
# STEP 13. Test 고성과 예측 및 실제성과 비교
# ============================================================

# Test 데이터 예측
test_prediction <- predict_employees(
  test_data,
  model,
  threshold
)

# Test 원본의 성과(0/1)를 실제성과로 변환
test_prediction$실제성과 <- ifelse(
  is.na(test_prediction$성과),
  NA_character_,
  ifelse(
    test_prediction$성과 == 1,
    "고성과",
    "저성과"
  )
)

# 실제성과와 예측분류 비교
test_prediction$예측정확여부 <- ifelse(
  is.na(test_prediction$예측성과),
  "예측 불가",
  ifelse(
    is.na(test_prediction$실제성과),
    "검증 대기",
    ifelse(
      test_prediction$실제성과 ==
        test_prediction$예측분류,
      "정확",
      "오분류"
    )
  )
)

# 최종 필드 순서
result_cols <- c(
  "응답자 번호",
  "고성과_예측확률",
  "예측성과",
  "예측분류",
  "실제성과",
  "예측정확여부"
)

other_cols <- setdiff(
  names(test_prediction),
  result_cols
)

test_prediction <- test_prediction[
  ,
  c(
    other_cols,
    result_cols
  ),
  drop = FALSE
]

# 실제 성과를 활용한 Test 성능 평가
test_eval <- evaluate_model(
  test_prediction,
  threshold
)

cat("\n===== Test 예측 완료 =====\n")

cat(
  "Test 예측 가능:",
  sum(
    !is.na(
      test_prediction$고성과_예측확률
    )
  ),
  "\n"
)

cat(
  "Test 실제성과 검증 가능:",
  sum(
    !is.na(test_prediction$성과) &
      !is.na(test_prediction$예측성과)
  ),
  "\n"
)


# ============================================================
# STEP 14. Train / Test 성능 비교
# ============================================================

metric_names <- c(
  "Accuracy",
  "Precision",
  "Sensitivity",
  "Specificity",
  "F1",
  "AUC"
)

comparison <- data.frame(
  지표 = metric_names,
  Train = as.numeric(
    train_eval$metrics[metric_names]
  ),
  Test = as.numeric(
    test_eval$metrics[metric_names]
  )
)

comparison$차이_Test_Train <-
  comparison$Test -
  comparison$Train

comparison$`결과 해석` <- vapply(
  seq_len(nrow(comparison)),
  function(i) {
    
    delta <-
      comparison$차이_Test_Train[i]
    
    if (is.na(delta)) {
      
      return(
        "비교 불가: 실제 성과 또는 계산값 확인 필요"
      )
    }
    
    if (abs(delta) < 0.001) {
      
      return(
        "Train과 Test 성능이 거의 동일"
      )
    }
    
    if (delta < 0) {
      
      return(
        paste0(
          "Test 성능이 Train보다 ",
          sprintf(
            "%.1f",
            abs(delta) * 100
          ),
          "%p 낮음"
        )
      )
    }
    
    paste0(
      "Test 성능이 Train보다 ",
      sprintf(
        "%.1f",
        delta * 100
      ),
      "%p 높음"
    )
  },
  character(1)
)


# ============================================================
# STEP 15. Excel 결과 통합
# ============================================================

wb <- openxlsx::createWorkbook()

add_sheet <- function(
    sheet_name,
    df
) {
  
  openxlsx::addWorksheet(
    wb,
    sheet_name
  )
  
  header_style <- openxlsx::createStyle(
    fontColour = "#FFFFFF",
    fgFill = "#243B55",
    textDecoration = "bold",
    halign = "center",
    valign = "center"
  )
  
  openxlsx::writeData(
    wb,
    sheet_name,
    df,
    headerStyle = header_style
  )
  
  openxlsx::freezePane(
    wb,
    sheet_name,
    firstRow = TRUE
  )
  
  openxlsx::setColWidths(
    wb,
    sheet_name,
    cols = seq_len(ncol(df)),
    widths = "auto"
  )
  
  if (nrow(df) > 0) {
    
    openxlsx::addFilter(
      wb,
      sheet_name,
      rows = 1,
      cols = seq_len(ncol(df))
    )
  }
}


# 01~09 분석 결과
add_sheet(
  "01_기술통계",
  descriptive
)

add_sheet(
  "02_범주형빈도",
  categorical_summary
)

add_sheet(
  "03_로지스틱회귀_OR",
  or_results
)

add_sheet(
  "04_다중공선성",
  vif_results
)

add_sheet(
  "05_모형적합도",
  model_fit
)

add_sheet(
  "06_혼동행렬",
  train_eval$confusion
)

add_sheet(
  "07_예측성능",
  train_eval$performance
)

add_sheet(
  "08_ROC데이터",
  roc_points
)

add_sheet(
  "09_직원별예측",
  train_prediction
)


# 분석 설정
analysis_info <- data.frame(
  항목 = c(
    "분석 방법",
    "종속변수",
    "고성과 기준",
    "결혼여부 기준집단",
    "성별 기준집단",
    "학력 기준집단",
    "학력 비교집단",
    "예측 분류 기준",
    "OR 신뢰구간",
    "Train 성능 평가",
    "Test 성능 평가",
    "Train 파일",
    "Test 파일",
    "분석 실행일"
  ),
  내용 = c(
    "Binary Logistic Regression",
    "성과",
    "성과 = 1",
    "M",
    "F",
    "1",
    "2 vs 1, 3 vs 1",
    as.character(threshold),
    "Wald 95% CI",
    "학습 데이터 (In-sample)",
    "별도 Test 데이터 (Holdout)",
    basename(train_path),
    basename(test_path),
    as.character(Sys.time())
  )
)

add_sheet(
  "10_분석설정",
  analysis_info
)


# ------------------------------------------------------------
# 11_Test직원별예측
# ------------------------------------------------------------

sheet_name <- "11_Test직원별예측"

add_sheet(
  sheet_name,
  test_prediction
)

# 실제성과 및 예측정확여부 열 위치
actual_col <- match(
  "실제성과",
  names(test_prediction)
)

pred_col <- match(
  "예측분류",
  names(test_prediction)
)

check_col <- match(
  "예측정확여부",
  names(test_prediction)
)

if (nrow(test_prediction) > 0) {
  
  # 실제성과와 예측분류를 Excel에서 자동 비교
  for (i in seq_len(
    nrow(test_prediction)
  )) {
    
    excel_row <- i + 1
    
    actual_cell <- paste0(
      openxlsx::int2col(actual_col),
      excel_row
    )
    
    pred_cell <- paste0(
      openxlsx::int2col(pred_col),
      excel_row
    )
    
    formula <- paste0(
      'IF(',
      pred_cell,
      '="","예측 불가",',
      'IF(',
      actual_cell,
      '="","검증 대기",',
      'IF(',
      actual_cell,
      '=',
      pred_cell,
      ',"정확","오분류")))'
    )
    
    openxlsx::writeFormula(
      wb,
      sheet = sheet_name,
      x = formula,
      startCol = check_col,
      startRow = excel_row
    )
  }
  
  # 실제성과 수정용 드롭다운
  openxlsx::dataValidation(
    wb,
    sheet = sheet_name,
    cols = actual_col,
    rows = 2:(
      nrow(test_prediction) + 1
    ),
    type = "list",
    value = '"고성과,저성과"'
  )
}


# 12~14 Test 성능 평가 및 비교
add_sheet(
  "12_Test예측성능",
  test_eval$performance
)

add_sheet(
  "13_Test혼동행렬",
  test_eval$confusion
)

add_sheet(
  "14_Train_Test비교",
  comparison
)


# ============================================================
# STEP 16. Excel 결과 저장
# ============================================================

output_file <- file.path(
  output_dir,
  paste0(
    tools::file_path_sans_ext(
      basename(train_path)
    ),
    "_Binary_Logistic_Results.xlsx"
  )
)

openxlsx::saveWorkbook(
  wb,
  output_file,
  overwrite = TRUE
)


# ============================================================
# STEP 17. 분석 결과 출력
# ============================================================

cat("\n====================================\n")
cat("직원 성과 이항 로지스틱 회귀분석 완료\n")
cat("====================================\n")

cat("\n[1] 회귀분석 및 OR\n")
print(or_results)

cat("\n[2] 다중공선성\n")
print(vif_results)

cat("\n[3] 모형 적합도\n")
print(model_fit)

cat("\n[4] Train 예측 성능\n")
print(train_eval$performance)

cat("\n[5] Test 예측 성능\n")
print(test_eval$performance)

cat("\n[6] Train / Test 비교\n")
print(comparison)

cat("\n[7] Test 직원별 예측 미리보기\n")

preview_cols <- c(
  "응답자 번호",
  "결혼여부",
  "고성과_예측확률",
  "예측성과",
  "예측분류",
  "실제성과",
  "예측정확여부"
)

print(
  head(
    test_prediction[
      ,
      preview_cols,
      drop = FALSE
    ],
    10
  )
)


# Train ROC 곡선
if (!is.null(train_eval$roc)) {
  
  plot(
    train_eval$roc,
    main = paste0(
      "Train ROC (AUC = ",
      round(
        train_eval$metrics["AUC"],
        3
      ),
      ")"
    ),
    legacy.axes = TRUE,
    lwd = 2
  )
  
  abline(
    a = 0,
    b = 1,
    lty = 2,
    col = "gray"
  )
}


# Test ROC 곡선
if (!is.null(test_eval$roc)) {
  
  plot(
    test_eval$roc,
    main = paste0(
      "Test ROC (AUC = ",
      round(
        test_eval$metrics["AUC"],
        3
      ),
      ")"
    ),
    legacy.axes = TRUE,
    lwd = 2
  )
  
  abline(
    a = 0,
    b = 1,
    lty = 2,
    col = "gray"
  )
}


cat("\n====================================\n")
cat("Excel 결과 저장 완료\n")
cat(output_file, "\n")
cat("====================================\n")
