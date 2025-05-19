\* Sourcing_with_Freight *\
Minimize
OBJ: F_Supplier_1 + F_Supplier_2 + F_Supplier_3 + F_Supplier_4 + F_Supplier_5
 + S_Supplier_1 + S_Supplier_2 + S_Supplier_3 + S_Supplier_4 + S_Supplier_5
 - reb_Supplier_1 - reb_Supplier_2 - reb_Supplier_3 - reb_Supplier_4
 - reb_Supplier_5
Subject To
Cap_Supplier_1_Bid_ID_1: x_Supplier_1_1 <= 1457
Cap_Supplier_1_Bid_ID_2: x_Supplier_1_2 <= 2422
Cap_Supplier_1_Bid_ID_3: x_Supplier_1_3 <= 6463
Cap_Supplier_1_Bid_ID_4: x_Supplier_1_4 <= 2358
Cap_Supplier_1_Bid_ID_5: x_Supplier_1_5 <= 2422
Cap_Supplier_2_Capacity_Group_Category_1: x_Supplier_2_1 + x_Supplier_2_3
 + x_Supplier_2_4 <= 8001
Cap_Supplier_2_Capacity_Group_Category_2: x_Supplier_2_2 <= 2501
Cap_Supplier_2_Capacity_Group_Category_3: x_Supplier_2_5 <= 501
Cap_Supplier_3_Bid_ID_1: x_Supplier_3_1 <= 1457
Cap_Supplier_3_Bid_ID_2: x_Supplier_3_2 <= 2422
Cap_Supplier_3_Bid_ID_3: x_Supplier_3_3 <= 6463
Cap_Supplier_3_Bid_ID_4: x_Supplier_3_4 <= 2358
Cap_Supplier_3_Bid_ID_5: x_Supplier_3_5 <= 2422
Cap_Supplier_4_Bid_ID_1: x_Supplier_4_1 <= 1457
Cap_Supplier_4_Bid_ID_2: x_Supplier_4_2 <= 2422
Cap_Supplier_4_Bid_ID_3: x_Supplier_4_3 <= 6463
Cap_Supplier_4_Bid_ID_4: x_Supplier_4_4 <= 2358
Cap_Supplier_4_Bid_ID_5: x_Supplier_4_5 <= 2422
Cap_Supplier_5_Description_Large_Item: x_Supplier_5_5 <= 5001
Cap_Supplier_5_Description_Medium_item: x_Supplier_5_2 <= 2501
Cap_Supplier_5_Description_Small_item: x_Supplier_5_1 + x_Supplier_5_3
 + x_Supplier_5_4 <= 9001
DDP_ON_Supplier_1_1: - 1000000000 bDDP_Supplier_1_1 + xSF_Supplier_1_1 <= 0
DDP_ON_Supplier_1_2: - 1000000000 bDDP_Supplier_1_2 + xSF_Supplier_1_2 <= 0
DDP_ON_Supplier_1_3: - 1000000000 bDDP_Supplier_1_3 + xSF_Supplier_1_3 <= 0
DDP_ON_Supplier_1_4: - 1000000000 bDDP_Supplier_1_4 + xSF_Supplier_1_4 <= 0
DDP_ON_Supplier_1_5: - 1000000000 bDDP_Supplier_1_5 + xSF_Supplier_1_5 <= 0
DDP_ON_Supplier_2_1: - 1000000000 bDDP_Supplier_2_1 + xSF_Supplier_2_1 <= 0
DDP_ON_Supplier_2_2: - 1000000000 bDDP_Supplier_2_2 + xSF_Supplier_2_2 <= 0
DDP_ON_Supplier_2_3: - 1000000000 bDDP_Supplier_2_3 + xSF_Supplier_2_3 <= 0
DDP_ON_Supplier_2_4: - 1000000000 bDDP_Supplier_2_4 + xSF_Supplier_2_4 <= 0
DDP_ON_Supplier_2_5: - 1000000000 bDDP_Supplier_2_5 + xSF_Supplier_2_5 <= 0
DDP_ON_Supplier_3_1: - 1000000000 bDDP_Supplier_3_1 + xSF_Supplier_3_1 <= 0
DDP_ON_Supplier_3_2: - 1000000000 bDDP_Supplier_3_2 + xSF_Supplier_3_2 <= 0
DDP_ON_Supplier_3_3: - 1000000000 bDDP_Supplier_3_3 + xSF_Supplier_3_3 <= 0
DDP_ON_Supplier_3_4: - 1000000000 bDDP_Supplier_3_4 + xSF_Supplier_3_4 <= 0
DDP_ON_Supplier_3_5: - 1000000000 bDDP_Supplier_3_5 + xSF_Supplier_3_5 <= 0
DDP_ON_Supplier_4_1: - 1000000000 bDDP_Supplier_4_1 + xSF_Supplier_4_1 <= 0
DDP_ON_Supplier_4_2: - 1000000000 bDDP_Supplier_4_2 + xSF_Supplier_4_2 <= 0
DDP_ON_Supplier_4_3: - 1000000000 bDDP_Supplier_4_3 + xSF_Supplier_4_3 <= 0
DDP_ON_Supplier_4_4: - 1000000000 bDDP_Supplier_4_4 + xSF_Supplier_4_4 <= 0
DDP_ON_Supplier_4_5: - 1000000000 bDDP_Supplier_4_5 + xSF_Supplier_4_5 <= 0
DDP_ON_Supplier_5_1: - 1000000000 bDDP_Supplier_5_1 + xSF_Supplier_5_1 <= 0
DDP_ON_Supplier_5_2: - 1000000000 bDDP_Supplier_5_2 + xSF_Supplier_5_2 <= 0
DDP_ON_Supplier_5_3: - 1000000000 bDDP_Supplier_5_3 + xSF_Supplier_5_3 <= 0
DDP_ON_Supplier_5_4: - 1000000000 bDDP_Supplier_5_4 + xSF_Supplier_5_4 <= 0
DDP_ON_Supplier_5_5: - 1000000000 bDDP_Supplier_5_5 + xSF_Supplier_5_5 <= 0
DMax_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_2 + x_Supplier_1_3
 + x_Supplier_1_4 + x_Supplier_1_5 + 1000000000 zd_Supplier_1_0 <= 1001000000
DMin_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_2 + x_Supplier_1_3
 + x_Supplier_1_4 + x_Supplier_1_5 - zd_Supplier_1_0 >= 0
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
F_Supplier_1: F_Supplier_1 - xKBX_Supplier_1_1 - xKBX_Supplier_1_2
 - xKBX_Supplier_1_3 - xKBX_Supplier_1_4 - xKBX_Supplier_1_5 = 0
F_Supplier_2: F_Supplier_2 - xKBX_Supplier_2_1 - xKBX_Supplier_2_2
 - xKBX_Supplier_2_3 - xKBX_Supplier_2_4 - xKBX_Supplier_2_5 = 0
F_Supplier_3: F_Supplier_3 - xKBX_Supplier_3_1 - xKBX_Supplier_3_2
 - xKBX_Supplier_3_3 - xKBX_Supplier_3_4 - xKBX_Supplier_3_5 = 0
F_Supplier_4: F_Supplier_4 - xKBX_Supplier_4_1 - xKBX_Supplier_4_2
 - xKBX_Supplier_4_3 - xKBX_Supplier_4_4 - xKBX_Supplier_4_5 = 0
F_Supplier_5: F_Supplier_5 - xKBX_Supplier_5_1 - xKBX_Supplier_5_2
 - xKBX_Supplier_5_3 - xKBX_Supplier_5_4 - xKBX_Supplier_5_5 = 0
Fixd_Supplier_2: d_Supplier_2 = 0
Fixd_Supplier_3: d_Supplier_3 = 0
Fixd_Supplier_4: d_Supplier_4 = 0
Fixd_Supplier_5: d_Supplier_5 = 0
Fixreb_Supplier_2: reb_Supplier_2 = 0
Fixreb_Supplier_3: reb_Supplier_3 = 0
Fixreb_Supplier_4: reb_Supplier_4 = 0
Fixreb_Supplier_5: reb_Supplier_5 = 0
KBX_ON_Supplier_1_1: 1000000000 bDDP_Supplier_1_1 + xKBX_Supplier_1_1
 <= 1000000000
KBX_ON_Supplier_1_2: 1000000000 bDDP_Supplier_1_2 + xKBX_Supplier_1_2
 <= 1000000000
KBX_ON_Supplier_1_3: 1000000000 bDDP_Supplier_1_3 + xKBX_Supplier_1_3
 <= 1000000000
KBX_ON_Supplier_1_4: 1000000000 bDDP_Supplier_1_4 + xKBX_Supplier_1_4
 <= 1000000000
KBX_ON_Supplier_1_5: 1000000000 bDDP_Supplier_1_5 + xKBX_Supplier_1_5
 <= 1000000000
KBX_ON_Supplier_2_1: 1000000000 bDDP_Supplier_2_1 + xKBX_Supplier_2_1
 <= 1000000000
KBX_ON_Supplier_2_2: 1000000000 bDDP_Supplier_2_2 + xKBX_Supplier_2_2
 <= 1000000000
KBX_ON_Supplier_2_3: 1000000000 bDDP_Supplier_2_3 + xKBX_Supplier_2_3
 <= 1000000000
KBX_ON_Supplier_2_4: 1000000000 bDDP_Supplier_2_4 + xKBX_Supplier_2_4
 <= 1000000000
KBX_ON_Supplier_2_5: 1000000000 bDDP_Supplier_2_5 + xKBX_Supplier_2_5
 <= 1000000000
KBX_ON_Supplier_3_1: 1000000000 bDDP_Supplier_3_1 + xKBX_Supplier_3_1
 <= 1000000000
KBX_ON_Supplier_3_2: 1000000000 bDDP_Supplier_3_2 + xKBX_Supplier_3_2
 <= 1000000000
KBX_ON_Supplier_3_3: 1000000000 bDDP_Supplier_3_3 + xKBX_Supplier_3_3
 <= 1000000000
KBX_ON_Supplier_3_4: 1000000000 bDDP_Supplier_3_4 + xKBX_Supplier_3_4
 <= 1000000000
KBX_ON_Supplier_3_5: 1000000000 bDDP_Supplier_3_5 + xKBX_Supplier_3_5
 <= 1000000000
KBX_ON_Supplier_4_1: 1000000000 bDDP_Supplier_4_1 + xKBX_Supplier_4_1
 <= 1000000000
KBX_ON_Supplier_4_2: 1000000000 bDDP_Supplier_4_2 + xKBX_Supplier_4_2
 <= 1000000000
KBX_ON_Supplier_4_3: 1000000000 bDDP_Supplier_4_3 + xKBX_Supplier_4_3
 <= 1000000000
KBX_ON_Supplier_4_4: 1000000000 bDDP_Supplier_4_4 + xKBX_Supplier_4_4
 <= 1000000000
KBX_ON_Supplier_4_5: 1000000000 bDDP_Supplier_4_5 + xKBX_Supplier_4_5
 <= 1000000000
KBX_ON_Supplier_5_1: 1000000000 bDDP_Supplier_5_1 + xKBX_Supplier_5_1
 <= 1000000000
KBX_ON_Supplier_5_2: 1000000000 bDDP_Supplier_5_2 + xKBX_Supplier_5_2
 <= 1000000000
KBX_ON_Supplier_5_3: 1000000000 bDDP_Supplier_5_3 + xKBX_Supplier_5_3
 <= 1000000000
KBX_ON_Supplier_5_4: 1000000000 bDDP_Supplier_5_4 + xKBX_Supplier_5_4
 <= 1000000000
KBX_ON_Supplier_5_5: 1000000000 bDDP_Supplier_5_5 + xKBX_Supplier_5_5
 <= 1000000000
LinkVol_Supplier_1_1: - xKBX_Supplier_1_1 - xSF_Supplier_1_1 + x_Supplier_1_1
 = 0
LinkVol_Supplier_1_2: - xKBX_Supplier_1_2 - xSF_Supplier_1_2 + x_Supplier_1_2
 = 0
LinkVol_Supplier_1_3: - xKBX_Supplier_1_3 - xSF_Supplier_1_3 + x_Supplier_1_3
 = 0
LinkVol_Supplier_1_4: - xKBX_Supplier_1_4 - xSF_Supplier_1_4 + x_Supplier_1_4
 = 0
LinkVol_Supplier_1_5: - xKBX_Supplier_1_5 - xSF_Supplier_1_5 + x_Supplier_1_5
 = 0
LinkVol_Supplier_2_1: - xKBX_Supplier_2_1 - xSF_Supplier_2_1 + x_Supplier_2_1
 = 0
LinkVol_Supplier_2_2: - xKBX_Supplier_2_2 - xSF_Supplier_2_2 + x_Supplier_2_2
 = 0
LinkVol_Supplier_2_3: - xKBX_Supplier_2_3 - xSF_Supplier_2_3 + x_Supplier_2_3
 = 0
LinkVol_Supplier_2_4: - xKBX_Supplier_2_4 - xSF_Supplier_2_4 + x_Supplier_2_4
 = 0
LinkVol_Supplier_2_5: - xKBX_Supplier_2_5 - xSF_Supplier_2_5 + x_Supplier_2_5
 = 0
LinkVol_Supplier_3_1: - xKBX_Supplier_3_1 - xSF_Supplier_3_1 + x_Supplier_3_1
 = 0
LinkVol_Supplier_3_2: - xKBX_Supplier_3_2 - xSF_Supplier_3_2 + x_Supplier_3_2
 = 0
LinkVol_Supplier_3_3: - xKBX_Supplier_3_3 - xSF_Supplier_3_3 + x_Supplier_3_3
 = 0
LinkVol_Supplier_3_4: - xKBX_Supplier_3_4 - xSF_Supplier_3_4 + x_Supplier_3_4
 = 0
LinkVol_Supplier_3_5: - xKBX_Supplier_3_5 - xSF_Supplier_3_5 + x_Supplier_3_5
 = 0
LinkVol_Supplier_4_1: - xKBX_Supplier_4_1 - xSF_Supplier_4_1 + x_Supplier_4_1
 = 0
LinkVol_Supplier_4_2: - xKBX_Supplier_4_2 - xSF_Supplier_4_2 + x_Supplier_4_2
 = 0
LinkVol_Supplier_4_3: - xKBX_Supplier_4_3 - xSF_Supplier_4_3 + x_Supplier_4_3
 = 0
LinkVol_Supplier_4_4: - xKBX_Supplier_4_4 - xSF_Supplier_4_4 + x_Supplier_4_4
 = 0
LinkVol_Supplier_4_5: - xKBX_Supplier_4_5 - xSF_Supplier_4_5 + x_Supplier_4_5
 = 0
LinkVol_Supplier_5_1: - xKBX_Supplier_5_1 - xSF_Supplier_5_1 + x_Supplier_5_1
 = 0
LinkVol_Supplier_5_2: - xKBX_Supplier_5_2 - xSF_Supplier_5_2 + x_Supplier_5_2
 = 0
LinkVol_Supplier_5_3: - xKBX_Supplier_5_3 - xSF_Supplier_5_3 + x_Supplier_5_3
 = 0
LinkVol_Supplier_5_4: - xKBX_Supplier_5_4 - xSF_Supplier_5_4 + x_Supplier_5_4
 = 0
LinkVol_Supplier_5_5: - xKBX_Supplier_5_5 - xSF_Supplier_5_5 + x_Supplier_5_5
 = 0
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
OneDisc_Supplier_1: - z_Supplier_1 + zd_Supplier_1_0 = 0
OneReb_Supplier_1: yr_Supplier_1_0 - z_Supplier_1 = 0
RMax_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_2 + x_Supplier_1_3
 + x_Supplier_1_4 + x_Supplier_1_5 + 1000000000 yr_Supplier_1_0 <= 1001000000
RMin_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_2 + x_Supplier_1_3
 + x_Supplier_1_4 + x_Supplier_1_5 - yr_Supplier_1_0 >= 0
S0_Supplier_1: S0_Supplier_1 - 34 xKBX_Supplier_1_1 - 15 xKBX_Supplier_1_2
 - 15 xKBX_Supplier_1_3 - 75 xKBX_Supplier_1_4 - 24 xKBX_Supplier_1_5
 - 36 xSF_Supplier_1_1 - 17 xSF_Supplier_1_2 - 17 xSF_Supplier_1_3
 - 77 xSF_Supplier_1_4 - 26 xSF_Supplier_1_5 = 0
S0_Supplier_2: S0_Supplier_2 - 14 xKBX_Supplier_2_1 - 15 xKBX_Supplier_2_2
 - 78 xKBX_Supplier_2_3 - 34 xKBX_Supplier_2_4 - 15 xKBX_Supplier_2_5
 - 16 xSF_Supplier_2_1 - 17 xSF_Supplier_2_2 - 80 xSF_Supplier_2_3
 - 36 xSF_Supplier_2_4 - 17 xSF_Supplier_2_5 = 0
S0_Supplier_3: S0_Supplier_3 - 75 xKBX_Supplier_3_1 - 25 xKBX_Supplier_3_2
 - 56 xKBX_Supplier_3_3 - 24 xKBX_Supplier_3_4 - 87 xKBX_Supplier_3_5
 - 77 xSF_Supplier_3_1 - 27 xSF_Supplier_3_2 - 58 xSF_Supplier_3_3
 - 26 xSF_Supplier_3_4 - 89 xSF_Supplier_3_5 = 0
S0_Supplier_4: S0_Supplier_4 - 93 xKBX_Supplier_4_1 - 24 xKBX_Supplier_4_2
 - 78 xKBX_Supplier_4_3 - 56 xKBX_Supplier_4_4 - 86 xKBX_Supplier_4_5
 - 95 xSF_Supplier_4_1 - 26 xSF_Supplier_4_2 - 80 xSF_Supplier_4_3
 - 58 xSF_Supplier_4_4 - 88 xSF_Supplier_4_5 = 0
S0_Supplier_5: S0_Supplier_5 - 76 xKBX_Supplier_5_1 - 32 xKBX_Supplier_5_2
 - 89 xKBX_Supplier_5_3 - 13 xKBX_Supplier_5_4 - 68 xKBX_Supplier_5_5
 - 78 xSF_Supplier_5_1 - 34 xSF_Supplier_5_2 - 91 xSF_Supplier_5_3
 - 15 xSF_Supplier_5_4 - 70 xSF_Supplier_5_5 = 0
Spend_Supplier_1: - S0_Supplier_1 + S_Supplier_1 + d_Supplier_1 = 0
Spend_Supplier_2: - S0_Supplier_2 + S_Supplier_2 + d_Supplier_2 = 0
Spend_Supplier_3: - S0_Supplier_3 + S_Supplier_3 + d_Supplier_3 = 0
Spend_Supplier_4: - S0_Supplier_4 + S_Supplier_4 + d_Supplier_4 = 0
Spend_Supplier_5: - S0_Supplier_5 + S_Supplier_5 + d_Supplier_5 = 0
SupCntUB_0: Y_sup_0_Supplier_1 + Y_sup_0_Supplier_2 + Y_sup_0_Supplier_3
 + Y_sup_0_Supplier_4 + Y_sup_0_Supplier_5 <= 2
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
Ylb_0_Supplier_1: - 1e-06 Y_sup_0_Supplier_1 + x_Supplier_1_1 + x_Supplier_1_2
 + x_Supplier_1_3 + x_Supplier_1_4 + x_Supplier_1_5 >= 0
Ylb_0_Supplier_2: - 1e-06 Y_sup_0_Supplier_2 + x_Supplier_2_1 + x_Supplier_2_2
 + x_Supplier_2_3 + x_Supplier_2_4 + x_Supplier_2_5 >= 0
Ylb_0_Supplier_3: - 1e-06 Y_sup_0_Supplier_3 + x_Supplier_3_1 + x_Supplier_3_2
 + x_Supplier_3_3 + x_Supplier_3_4 + x_Supplier_3_5 >= 0
Ylb_0_Supplier_4: - 1e-06 Y_sup_0_Supplier_4 + x_Supplier_4_1 + x_Supplier_4_2
 + x_Supplier_4_3 + x_Supplier_4_4 + x_Supplier_4_5 >= 0
Ylb_0_Supplier_5: - 1e-06 Y_sup_0_Supplier_5 + x_Supplier_5_1 + x_Supplier_5_2
 + x_Supplier_5_3 + x_Supplier_5_4 + x_Supplier_5_5 >= 0
Yub_0_Supplier_1: - 1000000000 Y_sup_0_Supplier_1 + x_Supplier_1_1
 + x_Supplier_1_2 + x_Supplier_1_3 + x_Supplier_1_4 + x_Supplier_1_5 <= 0
Yub_0_Supplier_2: - 1000000000 Y_sup_0_Supplier_2 + x_Supplier_2_1
 + x_Supplier_2_2 + x_Supplier_2_3 + x_Supplier_2_4 + x_Supplier_2_5 <= 0
Yub_0_Supplier_3: - 1000000000 Y_sup_0_Supplier_3 + x_Supplier_3_1
 + x_Supplier_3_2 + x_Supplier_3_3 + x_Supplier_3_4 + x_Supplier_3_5 <= 0
Yub_0_Supplier_4: - 1000000000 Y_sup_0_Supplier_4 + x_Supplier_4_1
 + x_Supplier_4_2 + x_Supplier_4_3 + x_Supplier_4_4 + x_Supplier_4_5 <= 0
Yub_0_Supplier_5: - 1000000000 Y_sup_0_Supplier_5 + x_Supplier_5_1
 + x_Supplier_5_2 + x_Supplier_5_3 + x_Supplier_5_4 + x_Supplier_5_5 <= 0
dLow_Supplier_1_0: - 0.02 S0_Supplier_1 + d_Supplier_1
 - 1406346 zd_Supplier_1_0 >= -1406346
dUp_Supplier_1_0: - 0.02 S0_Supplier_1 + d_Supplier_1
 + 1406346 zd_Supplier_1_0 <= 1406346
dZeroNoTier_Supplier_2: d_Supplier_2 = 0
dZeroNoTier_Supplier_3: d_Supplier_3 = 0
dZeroNoTier_Supplier_4: d_Supplier_4 = 0
dZeroNoTier_Supplier_5: d_Supplier_5 = 0
dZero_Supplier_1: d_Supplier_1 - 1406346 zd_Supplier_1_0 <= 0
rLow_Supplier_1_0: - 0.95 S_Supplier_1 + reb_Supplier_1
 - 1406346 yr_Supplier_1_0 >= -1406346
rUp_Supplier_1_0: - 0.95 S_Supplier_1 + reb_Supplier_1
 + 1406346 yr_Supplier_1_0 <= 1406346
rZeroNoTier_Supplier_2: reb_Supplier_2 = 0
rZeroNoTier_Supplier_3: reb_Supplier_3 = 0
rZeroNoTier_Supplier_4: reb_Supplier_4 = 0
rZeroNoTier_Supplier_5: reb_Supplier_5 = 0
rZero_Supplier_1: reb_Supplier_1 - 1406346 yr_Supplier_1_0 <= 0
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
Y_sup_0_Supplier_1
Y_sup_0_Supplier_2
Y_sup_0_Supplier_3
Y_sup_0_Supplier_4
Y_sup_0_Supplier_5
bDDP_Supplier_1_1
bDDP_Supplier_1_2
bDDP_Supplier_1_3
bDDP_Supplier_1_4
bDDP_Supplier_1_5
bDDP_Supplier_2_1
bDDP_Supplier_2_2
bDDP_Supplier_2_3
bDDP_Supplier_2_4
bDDP_Supplier_2_5
bDDP_Supplier_3_1
bDDP_Supplier_3_2
bDDP_Supplier_3_3
bDDP_Supplier_3_4
bDDP_Supplier_3_5
bDDP_Supplier_4_1
bDDP_Supplier_4_2
bDDP_Supplier_4_3
bDDP_Supplier_4_4
bDDP_Supplier_4_5
bDDP_Supplier_5_1
bDDP_Supplier_5_2
bDDP_Supplier_5_3
bDDP_Supplier_5_4
bDDP_Supplier_5_5
yr_Supplier_1_0
z_Supplier_1
z_Supplier_2
z_Supplier_3
z_Supplier_4
z_Supplier_5
zd_Supplier_1_0
End
