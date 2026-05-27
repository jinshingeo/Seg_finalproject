library(sf)
library(dplyr)

# RSG 패키지 로드
setwd("/Users/jin/홍교수님 수업/RSG_final/analysis")
devtools::load_all("..", quiet = TRUE)

# ── 데이터 로드 ────────────────────────────────────────────────────────────────
data_sf <- readRDS("data_sf.rds") %>%
  st_set_crs(5179) %>%
  filter(!is.na(저소득))           # 일원2동(강남구) 제외

cat("분석 대상 행정동 수:", nrow(data_sf), "\n\n")

# ── 분석 파라미터 ──────────────────────────────────────────────────────────────
# bandwidth: 행정동 평균 반경 ~1km 고려, 1500m 설정
BW <- 1500

# ══════════════════════════════════════════════════════════════════════════════
# 1단계: 사회적 고립 측정 (저소득 독거노인 vs 전체 주민)
# ══════════════════════════════════════════════════════════════════════════════
cat("=== 1단계: 사회적 고립 측정 ===\n")
cat("저소득 독거노인 vs 전체 주민 (bandwidth =", BW, "m)\n\n")

result_I_global <- RSG_I_Global(
  data      = data_sf,
  group_a   = "저소득",
  total     = "총인구",
  bandwidth = BW,
  verbose   = TRUE
)

result_I_local <- RSG_I_Local(
  data      = data_sf,
  group_a   = "저소득",
  total     = "총인구",
  bandwidth = BW,
  verbose   = FALSE
)

cat("\n[전역 고립지수]\n")
cat("RSG_I_Global =", round(result_I_global$index, 4), "\n")
cat("(0에 가까울수록 사회 전체와 잘 섞임, 1에 가까울수록 고립)\n\n")

# 고립도 상위 행정동
local_I <- data.frame(
  구동            = paste0(data_sf$구, " ", data_sf$동),
  Local_Isolation = result_I_local$index
)
top_I <- local_I[order(-local_I$Local_Isolation), ][1:10, ]
cat("[고립지수 상위 10개 행정동]\n")
print(top_I)

# ══════════════════════════════════════════════════════════════════════════════
# 2단계: 소득 기반 분리 측정 (저소득 vs 비교집단 독거노인)
# ══════════════════════════════════════════════════════════════════════════════
cat("\n=== 2단계: 소득 기반 분리 측정 ===\n")
cat("저소득 독거노인 vs 비교집단 독거노인 (bandwidth =", BW, "m)\n\n")

result_D_global <- RSG_D_Global(
  data      = data_sf,
  group_a   = "저소득",
  group_b   = "비교집단",
  total     = "독거_전체",
  bandwidth = BW,
  verbose   = TRUE
)

result_D_local <- RSG_D_Local(
  data      = data_sf,
  group_a   = "저소득",
  group_b   = "비교집단",
  total     = "독거_전체",
  bandwidth = BW,
  verbose   = FALSE
)

cat("\n[전역 상이지수]\n")
cat("RSG_D_Global =", round(result_D_global$index, 4), "\n")
cat("(0: 완전 균등 분포, 1: 완전 분리)\n\n")

# 분리도 상위 행정동
local_D <- data.frame(
  구동                = paste0(data_sf$구, " ", data_sf$동),
  Local_Dissimilarity = result_D_local$index
)
top_D <- local_D[order(-local_D$Local_Dissimilarity), ][1:10, ]
cat("[상이지수 상위 10개 행정동]\n")
print(top_D)

# ══════════════════════════════════════════════════════════════════════════════
# 보완: NSI (다중 bandwidth)
# ══════════════════════════════════════════════════════════════════════════════
cat("\n=== 보완: NSI (분리의 공간적 스케일) ===\n")
bw_list <- c(500, 1000, 1500, 2000, 3000)

result_NSI <- RSG_NSI(
  data      = data_sf,
  group_a   = "저소득",
  group_b   = "비교집단",
  total     = "독거_전체",
  bandwidth = bw_list,
  verbose   = FALSE
)

cat("비공간 NSI (기준값):", round(result_NSI$aspatial, 4), "\n")
cat("bandwidth별 NSI:\n")
nsi_df <- data.frame(bandwidth_m = bw_list, NSI = round(result_NSI$index, 4))
print(nsi_df)

# ══════════════════════════════════════════════════════════════════════════════
# 결과 저장
# ══════════════════════════════════════════════════════════════════════════════
results <- list(
  I_global  = result_I_global,
  I_local   = result_I_local,
  D_global  = result_D_global,
  D_local   = result_D_local,
  NSI       = result_NSI,
  data_sf   = data_sf
)
saveRDS(results, "results.rds")
cat("\nresults.rds 저장 완료\n")
