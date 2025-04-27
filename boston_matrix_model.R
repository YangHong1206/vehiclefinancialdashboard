#### Libraries ####
library(lubridate)
library(tidyverse)
library(readxl)
library(agridat)
library(writexl)
library(broom) #linear regression model summary
library(glue)
library(car)
library(tidyplots)
library(openxlsx)
library(reshape2)
library(data.table)
library(rmarkdown)
library(caret) #逐步回归-找最优的变量
library(gptstudio) #使用chatgpt
library(ggthemes)#使用tableau的颜色
library(esquisse) #use,esquisser 拖放绘图
library(ggThemeAssist) # 交互式主题编辑
library(colourpicker)
library(tsviz)#时间序列可视化
library(cols4all)#颜色版 c4a_gui()交互界面


#### setting path ####
setwd("D:/Chrome_download")



#### data—Sources are manually renewable ####

rawdata <- read.csv("历史订单明细脱敏版.csv")
pre_apr <- read.csv("预审订单明细.csv")
act_channel<- read_xlsx("渠道管理列表.xlsx")
all_channel <- read_xlsx("渠道管理列表.xlsx")
channel <- read_excel("渠道管理列表.xlsx")
visit_record <- read_xlsx("店面拜访.xlsx")
华东员工<- read_xlsx("2025-03-02华东区花名册.xlsx")
历史累积拜访 <- read.xlsx("累积拜访数据源2112_2203.xlsx")
目标额台<- read_xlsx("D:/personal doc/杨宏/分公司助理/日报/目标转化/福州消费融目标.xlsx")

#全维度预审进件生效----

单号生效日 <- select(rawdata,申请编号,合同生效日期,申请状态,产品类目) %>% 
  filter(.,申请状态 != "订单已取消")#选择成交的订单日期


pre_apr<- left_join(pre_apr,单号生效日,
                    join_by("申请编号"=="申请编号")) #将成交订单合并到预审表格中

#write_excel_csv(pre_apr,"df.csv")

全维度预审量 <- pre_apr %>% 
  mutate(预审月= format(预审提交时间,"%Y%m"),
         预审结果通过订单 = case_when(预审结果 == "自动通过" ~1,.default = 0),
         是否预审复议是 = case_when(是否预审复议 == "是" ~1,.default = 0),
         预审通过 = 预审结果通过订单+是否预审复议是
  ) %>% 
  
  group_by(预审月, 店面城市, 
           产品类别, 产品类目 ,产品名称, 合作项目,
           车辆类型,主品牌,
           金融经理 ,金融顾问,渠道编码,
           渠道二级科目名称, 店面主体) %>% 
  
  summarise(预审量 =length(申请编号),
            预审通过 = sum(预审结果通过订单))


#write_excel_csv(全维度预审量, file = "全维度预审量.csv")

全维度进件量 <- pre_apr %>% 
  mutate(进件月 = format(进件时间, "%Y%m"),
         最后审批通过量 = if_else(最后一次人工审批时间 > 最后一次自动审批时间,
                           case_when(最后一次人工审批结果=="通过"~1,.default = 0),
                           case_when(最后一次自动审批结果=="通过"~1,.default = 0))
  ) %>% 
  group_by(进件月,店面城市, 
           产品类别, 产品类目,产品名称, 合作项目,
           车辆类型,主品牌,
           金融经理 ,金融顾问,渠道编码,
           渠道二级科目名称, 店面主体 ) %>%
  summarise(进件量 = length(申请编号),
            进件通过 = sum(最后审批通过量))

#write_excel_csv(全维度进件量,"全维度进件量.csv")


全维度生效量 <- rawdata %>% 
  mutate(生效月 = format(合同生效日期, "%Y%m"),
         生效量 = case_when(!is.na(合同生效日期) & 申请状态 != "订单已取消"~ 1,
                         .default = 0),
         生效额 = case_when(!is.na(合同生效日期) & 申请状态 != "订单已取消"~ 融资金额,
                         .default = 0)
         
  ) %>% 
  
  group_by(生效月, 店面城市,
           产品类别, 产品类目, 产品方案名称, 合作项目,
           车辆类型, 品牌,
           金融经理名称, 金融顾问名称,渠道代码,
           渠道二级科目, 店面主体) %>% 
  
  summarise(生效量 = sum(生效量),
            生效额 = sum(生效额, na.rm = TRUE))

#write_excel_csv(全维度生效量,"全维度生效量.csv")

#合并表格

df<- full_join(x= 全维度预审量, y= 全维度进件量 , 
               join_by("预审月"== "进件月", 
                       "店面城市"=="店面城市",
                       "产品类别"=="产品类别",
                       "产品类目"=="产品类目",
                       "产品名称"=="产品名称",
                       "合作项目"=="合作项目",
                       "车辆类型"=="车辆类型",
                       "主品牌"=="主品牌",
                       "金融经理"=="金融经理",
                       "金融顾问"=="金融顾问",
                       "渠道二级科目名称"=="渠道二级科目名称",
                       "渠道编码"=="渠道编码",
                       "店面主体"=="店面主体"
               ))


df2<- full_join(x= df, y= 全维度生效量 ,   #全维度预审进件生效量
                join_by("预审月"== "生效月", 
                        "店面城市"=="店面城市",
                        "产品类别"=="产品类别",
                        "产品类目"=="产品类目",
                        "产品名称"=="产品方案名称",
                        "合作项目"=="合作项目",
                        "车辆类型"=="车辆类型",
                        "主品牌"=="品牌",
                        "金融经理"=="金融经理名称",
                        "金融顾问"=="金融顾问名称",
                        "渠道二级科目名称"=="渠道二级科目",
                        "渠道编码" == "渠道代码",
                        "店面主体"=="店面主体"
                ))



顾问Boston象限数据源 <- df2 %>% 
  group_by(预审月,金融顾问) %>% 
  summarise( 生效量 = sum(生效量, na.rm = T)) %>% 
  drop_na(预审月) %>% 
  group_by(预审月) %>% 
  mutate(percent = 生效量/ sum(生效量)) %>% 
  group_by(金融顾问) %>% #聚合顾问
  mutate(环比 = 生效量 / lag(生效量) - 1) %>% #计算顾问上个月环比 
  ungroup() #再解开聚合



# 计算渠道市场增长率与份额
channel_data <- rawdata %>%
  group_by(渠道编号) %>%
  summarise(
    Growth_Rate = (sum(Current_Sales) - sum(LastYear_Sales)) / sum(LastYear_Sales),
    Market_Share = sum(Current_Sales) / total_sales
  )

# 波士顿矩阵分类
boston_matrix <- channel_data %>%
  mutate(
    Quadrant = case_when(
      Growth_Rate > 0.1 & Market_Share > 0.15 ~ "Star",
      Growth_Rate < 0.1 & Market_Share > 0.15 ~ "Cash Cow",
      Growth_Rate > 0.1 & Market_Share < 0.15 ~ "Question Mark",
      TRUE ~ "Dog"
    )
  )

# 可视化（ggplot2输出）
ggplot(boston_matrix, aes(Market_Share, Growth_Rate, color=Quadrant)) +
  geom_point(size=4) +
  geom_hline(yintercept=0.1, linetype="dashed") +
  geom_vline(xintercept=0.15, linetype="dashed") +
  labs(title="渠道波士顿矩阵分析 - 福州分公司")