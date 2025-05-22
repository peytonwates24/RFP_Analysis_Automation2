\* Sourcing_with_Freight *\
Minimize
OBJ: F_Supplier_1 + S_Supplier_1 - reb_Supplier_1
Subject To
Cap_Supplier_1_Bid_ID_1: x_Supplier_1_1 <= 1457
DDP_ON_Supplier_1_1: - 1000000000 bDDP_Supplier_1_1 + xSF_Supplier_1_1 <= 0
DMax_Supplier_1_0: x_Supplier_1_1 + 1000000000 zd_Supplier_1_0 <= 1001000000
DMin_Supplier_1_0: x_Supplier_1_1 - zd_Supplier_1_0 >= 0
Demand_1: x_Supplier_1_1 = 1456
F_Supplier_1: F_Supplier_1 - xKBX_Supplier_1_1 = 0
KBX_ON_Supplier_1_1: 1000000000 bDDP_Supplier_1_1 + xKBX_Supplier_1_1
 <= 1000000000
LinkVol_Supplier_1_1: - xKBX_Supplier_1_1 - xSF_Supplier_1_1 + x_Supplier_1_1
 = 0
Link_Supplier_1_1: x_Supplier_1_1 - 1000000000 z_Supplier_1 <= 0
MinAward_Supplier_1: x_Supplier_1_1 - z_Supplier_1 >= 0
OneDisc_Supplier_1: - z_Supplier_1 + zd_Supplier_1_0 = 0
OneReb_Supplier_1: yr_Supplier_1_0 - z_Supplier_1 = 0
RMax_Supplier_1_0: x_Supplier_1_1 + 1000000000 yr_Supplier_1_0 <= 1001000000
RMin_Supplier_1_0: x_Supplier_1_1 - yr_Supplier_1_0 >= 0
S0_Supplier_1: S0_Supplier_1 - 34 xKBX_Supplier_1_1 - 36 xSF_Supplier_1_1 = 0
Spend_Supplier_1: - S0_Supplier_1 + S_Supplier_1 + d_Supplier_1 = 0
TrLB_1_Supplier_1: x_Supplier_1_1 >= 0
TrUB_1_Supplier_1: - 1456 T_1_Supplier_1 + x_Supplier_1_1 <= 0
dLow_Supplier_1_0: - 0.02 S0_Supplier_1 + d_Supplier_1 - 49538 zd_Supplier_1_0
 >= -49538
dUp_Supplier_1_0: - 0.02 S0_Supplier_1 + d_Supplier_1 + 49538 zd_Supplier_1_0
 <= 49538
dZero_Supplier_1: d_Supplier_1 - 49538 zd_Supplier_1_0 <= 0
rLow_Supplier_1_0: - 0.95 S_Supplier_1 + reb_Supplier_1
 - 49538 yr_Supplier_1_0 >= -49538
rUp_Supplier_1_0: - 0.95 S_Supplier_1 + reb_Supplier_1 + 49538 yr_Supplier_1_0
 <= 49538
rZero_Supplier_1: reb_Supplier_1 - 49538 yr_Supplier_1_0 <= 0
Binaries
T_1_Supplier_1
bDDP_Supplier_1_1
yr_Supplier_1_0
z_Supplier_1
zd_Supplier_1_0
End
