# =====================================================
# 1. 패키지
# =====================================================

# 최초 1회만 실행
#install.packages("readxl")
#install.packages("writexl")
#install.packages("dplyr")

library(readxl)
library(writexl)
library(dplyr)


# =====================================================
# 2. 데이터 불러오기
# =====================================================

data <- read_excel(file.choose())


# =====================================================
# 3. 불안 척도 생성
# =====================================================

data <- data %>%
  mutate(
    
    학습불안 = rowMeans(
      select(.,
             학습불안1,
             학습불안2,
             학습불안3),
      na.rm = TRUE
    ),
    
    직무대체불안 = rowMeans(
      select(.,
             직무대체불안1,
             직무대체불안2,
             직무대체불안3),
      na.rm = TRUE
    ),
    
    프라이버시침해불안 = rowMeans(
      select(.,
             프라이버시침해불안1,
             프라이버시침해불안2,
             프라이버시침해불안3),
      na.rm = TRUE
    ),
    
    연령그룹 = ifelse(
      연령 < 50,
      "비고령(50세 미만)",
      "고령(50세 이상)"
    )
  )


# =====================================================
# 4. 분석 함수
# =====================================================

analyze_anxiety <- function(data, variable_name) {
  
  x <- data[[variable_name]]
  
  temp <- data.frame(
    value = x,
    group = data$연령그룹
  )
  
  temp <- na.omit(temp)
  
  young <- temp$value[
    temp$group == "비고령(50세 미만)"
  ]
  
  old <- temp$value[
    temp$group == "고령(50세 이상)"
  ]
  
  # Welch 독립표본 t-test
  test <- t.test(
    young,
    old,
    var.equal = FALSE
  )
  
  # Cohen's d
  n1 <- length(young)
  n2 <- length(old)
  
  pooled_sd <- sqrt(
    ((n1 - 1) * var(young) +
       (n2 - 1) * var(old)) /
      (n1 + n2 - 2)
  )
  
  d <- (
    mean(young) - mean(old)
  ) / pooled_sd
  
  effect_size <- case_when(
    abs(d) < 0.2 ~ "매우 작은 효과",
    abs(d) < 0.5 ~ "작은효과",
    abs(d) < 0.8 ~ "중간",
    TRUE ~ "큼"
  )
  
  data.frame(
    변수 = variable_name,
    
    비고령_N = n1,
    비고령_평균 = mean(young),
    비고령_SD = sd(young),
    
    고령_N = n2,
    고령_평균 = mean(old),
    고령_SD = sd(old),
    
    평균차이 = mean(young) - mean(old),
    
    t = unname(test$statistic),
    df = unname(test$parameter),
    p = test$p.value,
    
    CI_하한 = test$conf.int[1],
    CI_상한 = test$conf.int[2],
    
    Cohens_d = d,
    효과크기 = effect_size
  )
}


# =====================================================
# 5. Welch t-test + Cohen's d
# =====================================================

results <- bind_rows(
  analyze_anxiety(data, "학습불안"),
  analyze_anxiety(data, "직무대체불안"),
  analyze_anxiety(data, "프라이버시침해불안")
)


# =====================================================
# 6. 소수점 정리 - 자리수 확인
# =====================================================

results_final <- results %>%
  mutate(
    비고령_평균 = round(비고령_평균, 3),
    비고령_SD = round(비고령_SD, 3),
    
    고령_평균 = round(고령_평균, 3),
    고령_SD = round(고령_SD, 3),
    
    평균차이 = round(평균차이, 3),
    
    t = round(t, 3),
    df = round(df, 2),
    p = round(p, 4),
    
    CI_하한 = round(CI_하한, 3),
    CI_상한 = round(CI_상한, 3),
    
    Cohens_d = round(Cohens_d, 3)
  )


# =====================================================
# 7. 기술통계
# =====================================================

summary_table <- data %>%
  group_by(연령그룹) %>%
  summarise(
    n = n(),
    
    학습불안_M =
      mean(학습불안, na.rm = TRUE),
    학습불안_SD =
      sd(학습불안, na.rm = TRUE),
    
    직무대체불안_M =
      mean(직무대체불안, na.rm = TRUE),
    직무대체불안_SD =
      sd(직무대체불안, na.rm = TRUE),
    
    프라이버시침해불안_M =
      mean(프라이버시침해불안, na.rm = TRUE),
    프라이버시침해불안_SD =
      sd(프라이버시침해불안, na.rm = TRUE)
  )


# =====================================================
# 8. 결과 확인
# =====================================================

print(results_final)


# =====================================================
# 9. Excel 저장
# =====================================================

write_xlsx(
  list(
    "Welch_t검정" = results_final,
    "기술통계" = summary_table,
    "분석데이터" = data
  ),
  "AI불안연령집단_Welch_t검정.xlsx"
)
