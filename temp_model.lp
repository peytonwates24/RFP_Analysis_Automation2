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
F_Supplier_1: F_Supplier_1 - 0.412162289871 xKBX_Supplier_1_1
 - 0.22235953793 xKBX_Supplier_1_2 - 0.512780804906 xKBX_Supplier_1_3
 - 0.394681495501 xKBX_Supplier_1_4 - 0.0753850747199 xKBX_Supplier_1_5 = 0
F_Supplier_2: F_Supplier_2 - 0.757107072844 xKBX_Supplier_2_1
 - 0.356630788891 xKBX_Supplier_2_2 - 0.858608241067 xKBX_Supplier_2_3
 - 0.917768682562 xKBX_Supplier_2_4 - 0.967175460043 xKBX_Supplier_2_5 = 0
F_Supplier_3: F_Supplier_3 - 0.241023405628 xKBX_Supplier_3_1
 - 0.454237724257 xKBX_Supplier_3_2 - 0.129107439223 xKBX_Supplier_3_3
 - 0.451698077422 xKBX_Supplier_3_4 - 0.45801791591 xKBX_Supplier_3_5 = 0
F_Supplier_4: F_Supplier_4 - 0.255919123062 xKBX_Supplier_4_1
 - 0.559049287841 xKBX_Supplier_4_2 - 0.479616530055 xKBX_Supplier_4_3
 - 0.593669220863 xKBX_Supplier_4_4 - 0.33835492882 xKBX_Supplier_4_5 = 0
F_Supplier_5: F_Supplier_5 - 0.182453952076 xKBX_Supplier_5_1
 - 0.409663732116 xKBX_Supplier_5_2 - 0.770440355959 xKBX_Supplier_5_3
 - 0.0145714073022 xKBX_Supplier_5_4 - 0.728726258075 xKBX_Supplier_5_5 = 0
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
RMax_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_2 + x_Supplier_1_3
 + x_Supplier_1_4 + x_Supplier_1_5 + 1000000000 yr_Supplier_1_0 <= 1001000000
RMin_Supplier_1_0: x_Supplier_1_1 + x_Supplier_1_2 + x_Supplier_1_3
 + x_Supplier_1_4 + x_Supplier_1_5 - yr_Supplier_1_0 >= 0
S0_Supplier_1: S0_Supplier_1 - 34 xKBX_Supplier_1_1 - 15 xKBX_Supplier_1_2
 - 15 xKBX_Supplier_1_3 - 75 xKBX_Supplier_1_4 - 24 xKBX_Supplier_1_5
 - 34.0453508742 xSF_Supplier_1_1 - 15.5713844108 xSF_Supplier_1_2
 - 15.6132383629 xSF_Supplier_1_3 - 75.3658246386 xSF_Supplier_1_4
 - 24.5552993592 xSF_Supplier_1_5 = 0
S0_Supplier_2: S0_Supplier_2 - 14 xKBX_Supplier_2_1 - 15 xKBX_Supplier_2_2
 - 78 xKBX_Supplier_2_3 - 34 xKBX_Supplier_2_4 - 15 xKBX_Supplier_2_5
 - 14.4996874379 xSF_Supplier_2_1 - 15.6610023131 xSF_Supplier_2_2
 - 78.583992345 xSF_Supplier_2_3 - 34.245663609 xSF_Supplier_2_4
 - 15.4376357648 xSF_Supplier_2_5 = 0
S0_Supplier_3: S0_Supplier_3 - 75 xKBX_Supplier_3_1 - 25 xKBX_Supplier_3_2
 - 56 xKBX_Supplier_3_3 - 24 xKBX_Supplier_3_4 - 87 xKBX_Supplier_3_5
 - 75.7343467169 xSF_Supplier_3_1 - 25.273169315 xSF_Supplier_3_2
 - 56.1815020183 xSF_Supplier_3_3 - 24.660791839 xSF_Supplier_3_4
 - 87.0636169186 xSF_Supplier_3_5 = 0
S0_Supplier_4: S0_Supplier_4 - 93 xKBX_Supplier_4_1 - 24 xKBX_Supplier_4_2
 - 78 xKBX_Supplier_4_3 - 56 xKBX_Supplier_4_4 - 86 xKBX_Supplier_4_5
 - 93.9042837523 xSF_Supplier_4_1 - 24.4606874043 xSF_Supplier_4_2
 - 78.4896300664 xSF_Supplier_4_3 - 56.3562030322 xSF_Supplier_4_4
 - 86.618544981 xSF_Supplier_4_5 = 0
S0_Supplier_5: S0_Supplier_5 - 76 xKBX_Supplier_5_1 - 32 xKBX_Supplier_5_2
 - 89 xKBX_Supplier_5_3 - 13 xKBX_Supplier_5_4 - 68 xKBX_Supplier_5_5
 - 76.022766062 xSF_Supplier_5_1 - 32.3245140214 xSF_Supplier_5_2
 - 89.583153403 xSF_Supplier_5_3 - 13.2006886512 xSF_Supplier_5_4
 - 68.7967611263 xSF_Supplier_5_5 = 0
Spend_Supplier_1: - S0_Supplier_1 + S_Supplier_1 + d_Supplier_1 = 0
Spend_Supplier_2: - S0_Supplier_2 + S_Supplier_2 + d_Supplier_2 = 0
Spend_Supplier_3: - S0_Supplier_3 + S_Supplier_3 + d_Supplier_3 = 0
Spend_Supplier_4: - S0_Supplier_4 + S_Supplier_4 + d_Supplier_4 = 0
Spend_Supplier_5: - S0_Supplier_5 + S_Supplier_5 + d_Supplier_5 = 0
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
