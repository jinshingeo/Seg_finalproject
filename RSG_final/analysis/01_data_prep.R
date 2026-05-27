library(sf)
library(dplyr)

# ── 경로 설정 ──────────────────────────────────────────────────────────────────
path_elderly  <- "../독거노인관련_0527/독거노인+현황(성별_동별)_2024.csv"
path_pop      <- "../독거노인관련_0527/등록인구(연령별_동별)_2024.csv"
path_shp      <- "/Users/jin/석사논문/통계지역경계/행정구역.shp"

# ── 1. 독거노인 현황 ────────────────────────────────────────────────────────────
# 헤더 5행 건너뛰고 읽기
# 컬럼: 시도 | 구 | 동 | 전체(계/남/여) | 기초수급(계/남/여) | 차상위(계/남/여) | 일반(계/남/여)
raw_elderly <- read.csv(path_elderly, header = FALSE, skip = 5,
                        fileEncoding = "UTF-8-BOM",
                        col.names = c("시도","구","동",
                                      "전체_계","전체_남","전체_여",
                                      "기초수급_계","기초수급_남","기초수급_여",
                                      "차상위_계","차상위_남","차상위_여",
                                      "일반_계","일반_남","일반_여"))

# 실제 행정동만 남기기 (소계 행 제거)
elderly <- raw_elderly %>%
  filter(동 != "소계", 구 != "소계") %>%
  # 구 이름 정규화
  mutate(구 = ifelse(구 == "동대문", "동대문구", 구)) %>%
  # 동 이름 오타 수정: 정능 → 정릉
  mutate(동 = gsub("정능", "정릉", 동)) %>%
  # "-" → 0 변환 후 numeric
  mutate(across(전체_계:일반_여, ~ as.numeric(ifelse(. == "-", "0", .)))) %>%
  # 강동구 상일1동+상일2동 → 상일동으로 합산
  mutate(동 = ifelse(구 == "강동구" & 동 %in% c("상일1동","상일2동"), "상일동", 동)) %>%
  group_by(구, 동) %>%
  summarise(across(전체_계:일반_여, sum, na.rm = TRUE), .groups = "drop") %>%
  # A안: 저소득 = 기초수급자만 / 비교집단 = 차상위 + 일반
  mutate(
    저소득   = 기초수급_계,
    비교집단 = 차상위_계 + 일반_계,
    독거_전체 = 전체_계
  ) %>%
  select(구, 동, 저소득, 비교집단, 독거_전체)

cat("독거노인 행정동 수:", nrow(elderly), "\n")

# ── 2. 등록인구 (전체 주민 — 1단계 고립지수 분모) ──────────────────────────────
raw_pop <- read.csv(path_pop, header = FALSE, skip = 2,
                    fileEncoding = "UTF-8-BOM",
                    col.names = c("구","동","항목","총인구",
                                  paste0("age_", 1:21)))

# 한국인, 동별 합계만 사용
pop <- raw_pop %>%
  filter(항목 == "한국인", 동 != "소계") %>%
  mutate(총인구 = as.numeric(gsub(",", "", 총인구))) %>%
  mutate(동 = ifelse(구 == "강동구" & 동 %in% c("상일1동","상일2동"), "상일동", 동)) %>%
  group_by(구, 동) %>%
  summarise(총인구 = sum(총인구, na.rm = TRUE), .groups = "drop")

cat("등록인구 행정동 수:", nrow(pop), "\n")

# ── 3. shapefile 로드 ───────────────────────────────────────────────────────────
shp <- st_read(path_shp, quiet = TRUE) %>%
  rename(구 = SIGUNGU_NM, 동 = ADM_NM) %>%
  select(구, 동, geometry)

# 동 이름 정규화: shapefile의 가운뎃점(·) → 마침표(.)
shp$동 <- gsub("·", ".", shp$동)

cat("shapefile 행정동 수:", nrow(shp), "\n")

# ── 4. 조인 ─────────────────────────────────────────────────────────────────────
# 독거노인 + 등록인구
data_joined <- elderly %>%
  left_join(pop, by = c("구", "동"))

# shapefile과 조인
data_sf <- shp %>%
  left_join(data_joined, by = c("구", "동"))

# 조인 결과 확인
n_missing <- sum(is.na(data_sf$저소득))
cat("조인 후 결측 행정동 수:", n_missing, "\n")

if (n_missing > 0) {
  cat("결측 행정동 목록:\n")
  print(data_sf[is.na(data_sf$저소득), c("구","동")])
}

cat("\n--- 데이터 요약 ---\n")
cat("서울 전체 저소득 독거노인:", sum(data_sf$저소득, na.rm=TRUE), "명\n")
cat("서울 전체 비교집단 독거노인:", sum(data_sf$비교집단, na.rm=TRUE), "명\n")
cat("서울 전체 총인구:", sum(data_sf$총인구, na.rm=TRUE), "명\n")

saveRDS(data_sf, "data_sf.rds")
cat("\ndata_sf.rds 저장 완료\n")
