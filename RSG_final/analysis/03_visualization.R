library(sf)
library(dplyr)
library(ggplot2)
library(scales)
library(RColorBrewer)
library(classInt)

setwd("/Users/jin/홍교수님 수업/RSG_final/analysis")

results  <- readRDS("results.rds")
data_sf  <- results$data_sf %>% st_set_crs(5179) %>% filter(!is.na(저소득))

data_sf$Local_Isolation     <- results$I_local$index
data_sf$Local_Dissimilarity <- results$D_local$index

# ── 공통 테마 ──────────────────────────────────────────────────────────────────
theme_map <- function(title, subtitle) {
  list(
    theme_void(),
    theme(
      plot.title    = element_text(size = 14, face = "bold", hjust = 0.5, margin = margin(b = 4)),
      plot.subtitle = element_text(size = 9,  hjust = 0.5, color = "grey40", margin = margin(b = 8)),
      legend.position   = "right",
      legend.title      = element_text(size = 9),
      legend.text       = element_text(size = 8),
      plot.background   = element_rect(fill = "white", color = NA),
      plot.margin       = margin(12, 12, 12, 12)
    ),
    labs(title = title, subtitle = subtitle)
  )
}

jenks_cut <- function(x, n = 5) {
  brks <- classIntervals(x, n = n, style = "jenks")$brks
  cut(x, breaks = brks, include.lowest = TRUE)
}

# ── Map 1: 국지적 고립지수 ─────────────────────────────────────────────────────
data_sf$iso_class <- jenks_cut(data_sf$Local_Isolation)

p1 <- ggplot(data_sf) +
  geom_sf(aes(fill = iso_class), color = "white", linewidth = 0.08) +
  scale_fill_brewer(
    palette = "YlOrRd",
    name    = "국지적\n고립지수",
    labels  = c("낮음", "", "중간", "", "높음"),
    na.value = "grey80"
  ) +
  theme_map(
    title    = "1단계: 사회적 고립 (RSG_I_Local)",
    subtitle = "저소득 독거노인 vs 전체 주민 | 전역 고립지수 = 0.0174 | bandwidth = 1,500 m"
  )

ggsave("map_isolation.png", p1, width = 9, height = 8, dpi = 200)
cat("map_isolation.png 저장\n")

# ── Map 2: 국지적 상이지수 ─────────────────────────────────────────────────────
data_sf$diss_class <- jenks_cut(data_sf$Local_Dissimilarity)

p2 <- ggplot(data_sf) +
  geom_sf(aes(fill = diss_class), color = "white", linewidth = 0.08) +
  scale_fill_brewer(
    palette = "PuRd",
    name    = "국지적\n상이지수",
    labels  = c("낮음", "", "중간", "", "높음"),
    na.value = "grey80"
  ) +
  theme_map(
    title    = "2단계: 소득 기반 분리 (RSG_D_Local)",
    subtitle = "저소득 vs 비교집단 독거노인 | 전역 상이지수 = 0.0866 | bandwidth = 1,500 m"
  )

ggsave("map_dissimilarity.png", p2, width = 9, height = 8, dpi = 200)
cat("map_dissimilarity.png 저장\n")

# ── Map 3: 두 지도 나란히 (PNG 합치기) ────────────────────────────────────────
png("map_combined.png", width = 1800, height = 800, res = 150)
gridlayout <- rbind(c(1, 2))
p1_grob <- ggplotGrob(p1)
p2_grob <- ggplotGrob(p2)
grid::grid.newpage()
grid::pushViewport(grid::viewport(layout = grid::grid.layout(1, 2)))
grid::pushViewport(grid::viewport(layout.pos.row = 1, layout.pos.col = 1))
grid::grid.draw(p1_grob)
grid::popViewport()
grid::pushViewport(grid::viewport(layout.pos.row = 1, layout.pos.col = 2))
grid::grid.draw(p2_grob)
grid::popViewport()
dev.off()
cat("map_combined.png 저장\n")

# ── NSI 차트 ──────────────────────────────────────────────────────────────────
nsi_df <- data.frame(
  bandwidth_m = c(500, 1000, 1500, 2000, 3000),
  NSI         = results$NSI$index
)
aspatial_nsi <- results$NSI$aspatial

p3 <- ggplot(nsi_df, aes(x = bandwidth_m, y = NSI)) +
  geom_hline(yintercept = aspatial_nsi, linetype = "dashed", color = "grey60", linewidth = 0.8) +
  annotate("text", x = 500, y = aspatial_nsi * 1.06,
           label = paste0("비공간 NSI (기준값) = ", round(aspatial_nsi, 4)),
           hjust = 0, size = 3.5, color = "grey50") +
  geom_line(color = "#9E0142", linewidth = 1.3) +
  geom_point(color = "#9E0142", size = 3.5) +
  geom_text(aes(label = round(NSI, 4)), vjust = -1, size = 3.2, color = "#9E0142") +
  scale_x_continuous(
    breaks = nsi_df$bandwidth_m,
    labels = paste0(nsi_df$bandwidth_m, " m")
  ) +
  scale_y_continuous(limits = c(0, aspatial_nsi * 1.35)) +
  labs(
    title    = "보완: NSI (근린 정렬 지수)",
    subtitle = "bandwidth가 넓어질수록 NSI 감소 → 소득 기반 분리는 광역보다 근린 수준에 집중",
    x = "Bandwidth", y = "NSI"
  ) +
  theme_minimal(base_size = 12) +
  theme(
    plot.title      = element_text(face = "bold"),
    plot.subtitle   = element_text(color = "grey40", size = 9),
    panel.grid.minor = element_blank(),
    plot.background = element_rect(fill = "white", color = NA),
    plot.margin     = margin(12, 16, 12, 12)
  )

ggsave("chart_nsi.png", p3, width = 8, height = 5, dpi = 200)
cat("chart_nsi.png 저장\n")

cat("\n=== 시각화 완료 ===\n")
cat("출력 파일:\n")
cat("  map_isolation.png     — 1단계 국지적 고립지수 지도\n")
cat("  map_dissimilarity.png — 2단계 국지적 상이지수 지도\n")
cat("  map_combined.png      — 두 지도 나란히\n")
cat("  chart_nsi.png         — NSI bandwidth 차트\n")
