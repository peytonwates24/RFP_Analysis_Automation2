\* Sourcing_with_Freight *\
Minimize
OBJ: S_Supplier_1 + S_Supplier_2 + S_Supplier_3 + S_Supplier_4 + S_Supplier_5
 - reb_Supplier_1 - reb_Supplier_2 - reb_Supplier_3 - reb_Supplier_4
 - reb_Supplier_5
Subject To
Cap_Supplier_1_Bid_ID_1: x_Supplier_1_1 <= 100000000000
Cap_Supplier_1_Bid_ID_2: x_Supplier_1_2 <= 100000000000
Cap_Supplier_1_Bid_ID_3: x_Supplier_1_3 <= 100000000000
Cap_Supplier_1_Bid_ID_4: x_Supplier_1_4 <= 100000000000
Cap_Supplier_1_Bid_ID_5: x_Supplier_1_5 <= 100000000000
Cap_Supplier_2_Bid_ID_1: x_Supplier_2_1 <= 100000000000
Cap_Supplier_2_Bid_ID_2: x_Supplier_2_2 <= 100000000000
Cap_Supplier_2_Bid_ID_3: x_Supplier_2_3 <= 100000000000
Cap_Supplier_2_Bid_ID_4: x_Supplier_2_4 <= 100000000000
Cap_Supplier_2_Bid_ID_5: x_Supplier_2_5 <= 100000000000
Cap_Supplier_3_Bid_ID_1: x_Supplier_3_1 <= 100000000000
Cap_Supplier_3_Bid_ID_2: x_Supplier_3_2 <= 100000000000
Cap_Supplier_3_Bid_ID_3: x_Supplier_3_3 <= 100000000000
Cap_Supplier_3_Bid_ID_4: x_Supplier_3_4 <= 100000000000
Cap_Supplier_3_Bid_ID_5: x_Supplier_3_5 <= 100000000000
Cap_Supplier_4_Bid_ID_1: x_Supplier_4_1 <= 100000000000
Cap_Supplier_4_Bid_ID_2: x_Supplier_4_2 <= 100000000000
Cap_Supplier_4_Bid_ID_3: x_Supplier_4_3 <= 100000000000
Cap_Supplier_4_Bid_ID_4: x_Supplier_4_4 <= 100000000000
Cap_Supplier_4_Bid_ID_5: x_Supplier_4_5 <= 100000000000
DMax_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_3 + x_Supplier_1_4
 + 1000000000 zd_Supplier_1_0 <= 1040000000
DMin_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_3 + x_Supplier_1_4
 - 20000000 zd_Supplier_1_0 >= 0
Demand_1: x_Supplier_1_1 + x_Supplier_2_1 + x_Supplier_3_1 + x_Supplier_4_1
 + x_Supplier_5_1 = 1456
Demand_2: x_Supplier_1_2 + x_Supplier_2_2 + x_Supplier_3_2 + x_Supplier_4_2
 + x_Supplier_5_2 = 2421
Demand_3: x_Supplier_1_3 + x_Supplier_2_3 + x_Supplier_3_3 + x_Supplier_4_3
 + x_Supplier_5_3 = 6462
Demand_4: x_Supplier_1_4 + x_Supplier_2_4 + x_Supplier_3_4 + x_Supplier_4_4
 + x_Supplier_5_4 = 2357
Demand_5: x_Supplier_1_5 + x_Supplier_2_5 + x_Supplier_3_5 + x_Supplier_4_5
 + x_Supplier_5_5 = 2421
DisableDisc_Supplier_1_0: zd_Supplier_1_0 = 0
F_Supplier_1: F_Supplier_1 - 3 x_Supplier_1_1 - 3 x_Supplier_1_2
 - 3 x_Supplier_1_3 - 3 x_Supplier_1_4 - 3 x_Supplier_1_5 = 0
F_Supplier_2: F_Supplier_2 - 3 x_Supplier_2_1 - 3 x_Supplier_2_2
 - 3 x_Supplier_2_3 - 3 x_Supplier_2_4 - 3 x_Supplier_2_5 = 0
F_Supplier_3: F_Supplier_3 - 3 x_Supplier_3_1 - 3 x_Supplier_3_2
 - 3 x_Supplier_3_3 - 3 x_Supplier_3_4 - 3 x_Supplier_3_5 = 0
F_Supplier_4: F_Supplier_4 - 3 x_Supplier_4_1 - 3 x_Supplier_4_2
 - 3 x_Supplier_4_3 - 3 x_Supplier_4_4 - 3 x_Supplier_4_5 = 0
F_Supplier_5: F_Supplier_5 - 3 x_Supplier_5_1 - 3 x_Supplier_5_2
 - 3 x_Supplier_5_3 - 3 x_Supplier_5_4 - 3 x_Supplier_5_5 = 0
Fixd_Supplier_2: d_Supplier_2 = 0
Fixd_Supplier_3: d_Supplier_3 = 0
Fixd_Supplier_4: d_Supplier_4 = 0
Fixd_Supplier_5: d_Supplier_5 = 0
Fixreb_Supplier_2: reb_Supplier_2 = 0
Fixreb_Supplier_3: reb_Supplier_3 = 0
Fixreb_Supplier_4: reb_Supplier_4 = 0
Fixreb_Supplier_5: reb_Supplier_5 = 0
Link_Supplier_1_1: x_Supplier_1_1 - 1000000000 z_Supplier_1 <= 0
Link_Supplier_1_2: x_Supplier_1_2 - 1000000000 z_Supplier_1 <= 0
Link_Supplier_1_3: x_Supplier_1_3 - 1000000000 z_Supplier_1 <= 0
Link_Supplier_1_4: x_Supplier_1_4 - 1000000000 z_Supplier_1 <= 0
Link_Supplier_1_5: x_Supplier_1_5 - 1000000000 z_Supplier_1 <= 0
Link_Supplier_2_1: x_Supplier_2_1 - 1000000000 z_Supplier_2 <= 0
Link_Supplier_2_2: x_Supplier_2_2 - 1000000000 z_Supplier_2 <= 0
Link_Supplier_2_3: x_Supplier_2_3 - 1000000000 z_Supplier_2 <= 0
Link_Supplier_2_4: x_Supplier_2_4 - 1000000000 z_Supplier_2 <= 0
Link_Supplier_2_5: x_Supplier_2_5 - 1000000000 z_Supplier_2 <= 0
Link_Supplier_3_1: x_Supplier_3_1 - 1000000000 z_Supplier_3 <= 0
Link_Supplier_3_2: x_Supplier_3_2 - 1000000000 z_Supplier_3 <= 0
Link_Supplier_3_3: x_Supplier_3_3 - 1000000000 z_Supplier_3 <= 0
Link_Supplier_3_4: x_Supplier_3_4 - 1000000000 z_Supplier_3 <= 0
Link_Supplier_3_5: x_Supplier_3_5 - 1000000000 z_Supplier_3 <= 0
Link_Supplier_4_1: x_Supplier_4_1 - 1000000000 z_Supplier_4 <= 0
Link_Supplier_4_2: x_Supplier_4_2 - 1000000000 z_Supplier_4 <= 0
Link_Supplier_4_3: x_Supplier_4_3 - 1000000000 z_Supplier_4 <= 0
Link_Supplier_4_4: x_Supplier_4_4 - 1000000000 z_Supplier_4 <= 0
Link_Supplier_4_5: x_Supplier_4_5 - 1000000000 z_Supplier_4 <= 0
Link_Supplier_5_1: x_Supplier_5_1 - 1000000000 z_Supplier_5 <= 0
Link_Supplier_5_2: x_Supplier_5_2 - 1000000000 z_Supplier_5 <= 0
Link_Supplier_5_3: x_Supplier_5_3 - 1000000000 z_Supplier_5 <= 0
Link_Supplier_5_4: x_Supplier_5_4 - 1000000000 z_Supplier_5 <= 0
Link_Supplier_5_5: x_Supplier_5_5 - 1000000000 z_Supplier_5 <= 0
MinAward_Supplier_1: x_Supplier_1_1 + x_Supplier_1_2 + x_Supplier_1_3
 + x_Supplier_1_4 + x_Supplier_1_5 - z_Supplier_1 >= 0
MinAward_Supplier_2: x_Supplier_2_1 + x_Supplier_2_2 + x_Supplier_2_3
 + x_Supplier_2_4 + x_Supplier_2_5 - z_Supplier_2 >= 0
MinAward_Supplier_3: x_Supplier_3_1 + x_Supplier_3_2 + x_Supplier_3_3
 + x_Supplier_3_4 + x_Supplier_3_5 - z_Supplier_3 >= 0
MinAward_Supplier_4: x_Supplier_4_1 + x_Supplier_4_2 + x_Supplier_4_3
 + x_Supplier_4_4 + x_Supplier_4_5 - z_Supplier_4 >= 0
MinAward_Supplier_5: x_Supplier_5_1 + x_Supplier_5_2 + x_Supplier_5_3
 + x_Supplier_5_4 + x_Supplier_5_5 - z_Supplier_5 >= 0
NoDisc_Supplier_1: zd_Supplier_1_0 = 0
OneReb_Supplier_1: yr_Supplier_1_0 = 1
PctAggEQLB_0: - x_Supplier_1_1 - x_Supplier_1_2 - x_Supplier_1_3
 + 0 x_Supplier_1_4 - x_Supplier_1_5 + 0 x_Supplier_2_1 - x_Supplier_2_2
 + 0 x_Supplier_2_3 - x_Supplier_2_4 - x_Supplier_2_5 - x_Supplier_3_1
 + 0 x_Supplier_3_2 - x_Supplier_3_3 - x_Supplier_3_4 + 0 x_Supplier_3_5
 - x_Supplier_4_1 - x_Supplier_4_2 - x_Supplier_4_3 - x_Supplier_4_4
 - x_Supplier_4_5 - x_Supplier_5_1 - x_Supplier_5_2 - x_Supplier_5_3
 - x_Supplier_5_4 - x_Supplier_5_5 >= 0
PctAggEQUB_0: - x_Supplier_1_1 - x_Supplier_1_2 - x_Supplier_1_3
 + 0 x_Supplier_1_4 - x_Supplier_1_5 + 0 x_Supplier_2_1 - x_Supplier_2_2
 + 0 x_Supplier_2_3 - x_Supplier_2_4 - x_Supplier_2_5 - x_Supplier_3_1
 + 0 x_Supplier_3_2 - x_Supplier_3_3 - x_Supplier_3_4 + 0 x_Supplier_3_5
 - x_Supplier_4_1 - x_Supplier_4_2 - x_Supplier_4_3 - x_Supplier_4_4
 - x_Supplier_4_5 - x_Supplier_5_1 - x_Supplier_5_2 - x_Supplier_5_3
 - x_Supplier_5_4 - x_Supplier_5_5 <= 0
RMax_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_3 + x_Supplier_1_4
 + 1000000000 yr_Supplier_1_0 <= 1000002000
RMin_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_3 + x_Supplier_1_4 >= 0
S0_Supplier_1: S0_Supplier_1 - 34 x_Supplier_1_1 - 15 x_Supplier_1_2
 - 15 x_Supplier_1_3 - 75 x_Supplier_1_4 - 24 x_Supplier_1_5 = 0
S0_Supplier_2: S0_Supplier_2 - 14 x_Supplier_2_1 - 15 x_Supplier_2_2
 - 78 x_Supplier_2_3 - 34 x_Supplier_2_4 - 15 x_Supplier_2_5 = 0
S0_Supplier_3: S0_Supplier_3 - 75 x_Supplier_3_1 - 25 x_Supplier_3_2
 - 56 x_Supplier_3_3 - 24 x_Supplier_3_4 - 87 x_Supplier_3_5 = 0
S0_Supplier_4: S0_Supplier_4 - 93 x_Supplier_4_1 - 24 x_Supplier_4_2
 - 78 x_Supplier_4_3 - 56 x_Supplier_4_4 - 86 x_Supplier_4_5 = 0
S0_Supplier_5: S0_Supplier_5 - 76 x_Supplier_5_1 - 32 x_Supplier_5_2
 - 89 x_Supplier_5_3 - 13 x_Supplier_5_4 - 68 x_Supplier_5_5 = 0
Spend_Supplier_1: - F_Supplier_1 - S0_Supplier_1 + S_Supplier_1 + d_Supplier_1
 = 0
Spend_Supplier_2: - F_Supplier_2 - S0_Supplier_2 + S_Supplier_2 + d_Supplier_2
 = 0
Spend_Supplier_3: - F_Supplier_3 - S0_Supplier_3 + S_Supplier_3 + d_Supplier_3
 = 0
Spend_Supplier_4: - F_Supplier_4 - S0_Supplier_4 + S_Supplier_4 + d_Supplier_4
 = 0
Spend_Supplier_5: - F_Supplier_5 - S0_Supplier_5 + S_Supplier_5 + d_Supplier_5
 = 0
TrLB_1_Supplier_1: x_Supplier_1_1 >= 0
TrLB_1_Supplier_3: x_Supplier_3_1 >= 0
TrLB_1_Supplier_4: x_Supplier_4_1 >= 0
TrLB_1_Supplier_5: x_Supplier_5_1 >= 0
TrLB_2_Supplier_1: x_Supplier_1_2 >= 0
TrLB_2_Supplier_2: x_Supplier_2_2 >= 0
TrLB_2_Supplier_4: x_Supplier_4_2 >= 0
TrLB_2_Supplier_5: x_Supplier_5_2 >= 0
TrLB_3_Supplier_1: x_Supplier_1_3 >= 0
TrLB_3_Supplier_3: x_Supplier_3_3 >= 0
TrLB_3_Supplier_4: x_Supplier_4_3 >= 0
TrLB_3_Supplier_5: x_Supplier_5_3 >= 0
TrLB_4_Supplier_2: x_Supplier_2_4 >= 0
TrLB_4_Supplier_3: x_Supplier_3_4 >= 0
TrLB_4_Supplier_4: x_Supplier_4_4 >= 0
TrLB_4_Supplier_5: x_Supplier_5_4 >= 0
TrLB_5_Supplier_1: x_Supplier_1_5 >= 0
TrLB_5_Supplier_2: x_Supplier_2_5 >= 0
TrLB_5_Supplier_4: x_Supplier_4_5 >= 0
TrLB_5_Supplier_5: x_Supplier_5_5 >= 0
TrUB_1_Supplier_1: - 1456 T_1_Supplier_1 + x_Supplier_1_1 <= 0
TrUB_1_Supplier_3: - 1456 T_1_Supplier_3 + x_Supplier_3_1 <= 0
TrUB_1_Supplier_4: - 1456 T_1_Supplier_4 + x_Supplier_4_1 <= 0
TrUB_1_Supplier_5: - 1456 T_1_Supplier_5 + x_Supplier_5_1 <= 0
TrUB_2_Supplier_1: - 2421 T_2_Supplier_1 + x_Supplier_1_2 <= 0
TrUB_2_Supplier_2: - 2421 T_2_Supplier_2 + x_Supplier_2_2 <= 0
TrUB_2_Supplier_4: - 2421 T_2_Supplier_4 + x_Supplier_4_2 <= 0
TrUB_2_Supplier_5: - 2421 T_2_Supplier_5 + x_Supplier_5_2 <= 0
TrUB_3_Supplier_1: - 6462 T_3_Supplier_1 + x_Supplier_1_3 <= 0
TrUB_3_Supplier_3: - 6462 T_3_Supplier_3 + x_Supplier_3_3 <= 0
TrUB_3_Supplier_4: - 6462 T_3_Supplier_4 + x_Supplier_4_3 <= 0
TrUB_3_Supplier_5: - 6462 T_3_Supplier_5 + x_Supplier_5_3 <= 0
TrUB_4_Supplier_2: - 2357 T_4_Supplier_2 + x_Supplier_2_4 <= 0
TrUB_4_Supplier_3: - 2357 T_4_Supplier_3 + x_Supplier_3_4 <= 0
TrUB_4_Supplier_4: - 2357 T_4_Supplier_4 + x_Supplier_4_4 <= 0
TrUB_4_Supplier_5: - 2357 T_4_Supplier_5 + x_Supplier_5_4 <= 0
TrUB_5_Supplier_1: - 2421 T_5_Supplier_1 + x_Supplier_1_5 <= 0
TrUB_5_Supplier_2: - 2421 T_5_Supplier_2 + x_Supplier_2_5 <= 0
TrUB_5_Supplier_4: - 2421 T_5_Supplier_4 + x_Supplier_4_5 <= 0
TrUB_5_Supplier_5: - 2421 T_5_Supplier_5 + x_Supplier_5_5 <= 0
V_Supplier_1: V_Supplier_1 - x_Supplier_1_1 - x_Supplier_1_2 - x_Supplier_1_3
 - x_Supplier_1_4 - x_Supplier_1_5 = 0
V_Supplier_2: V_Supplier_2 - x_Supplier_2_1 - x_Supplier_2_2 - x_Supplier_2_3
 - x_Supplier_2_4 - x_Supplier_2_5 = 0
V_Supplier_3: V_Supplier_3 - x_Supplier_3_1 - x_Supplier_3_2 - x_Supplier_3_3
 - x_Supplier_3_4 - x_Supplier_3_5 = 0
V_Supplier_4: V_Supplier_4 - x_Supplier_4_1 - x_Supplier_4_2 - x_Supplier_4_3
 - x_Supplier_4_4 - x_Supplier_4_5 = 0
V_Supplier_5: V_Supplier_5 - x_Supplier_5_1 - x_Supplier_5_2 - x_Supplier_5_3
 - x_Supplier_5_4 - x_Supplier_5_5 = 0
dLow_Supplier_1_0: - 0.05 S0_Supplier_1 + d_Supplier_1
 - 4.65e+13 zd_Supplier_1_0 >= -4.65e+13
dUp_Supplier_1_0: - 0.05 S0_Supplier_1 + d_Supplier_1
 + 4.65e+13 zd_Supplier_1_0 <= 4.65e+13
rLow_Supplier_1_0: - 0.05 S_Supplier_1 + reb_Supplier_1
 - 4.65e+13 yr_Supplier_1_0 >= -4.65e+13
rUp_Supplier_1_0: - 0.05 S_Supplier_1 + reb_Supplier_1
 + 4.65e+13 yr_Supplier_1_0 <= 4.65e+13
Binaries
T_1_Supplier_1
T_1_Supplier_3
T_1_Supplier_4
T_1_Supplier_5
T_2_Supplier_1
T_2_Supplier_2
T_2_Supplier_4
T_2_Supplier_5
T_3_Supplier_1
T_3_Supplier_3
T_3_Supplier_4
T_3_Supplier_5
T_4_Supplier_2
T_4_Supplier_3
T_4_Supplier_4
T_4_Supplier_5
T_5_Supplier_1
T_5_Supplier_2
T_5_Supplier_4
T_5_Supplier_5
yr_Supplier_1_0
z_Supplier_1
z_Supplier_2
z_Supplier_3
z_Supplier_4
z_Supplier_5
zd_Supplier_1_0
End
