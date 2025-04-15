\* Sourcing_with_MultiTier_Rebates_Discounts *\
Minimize
OBJ: S_Supplier_1 + S_Supplier_2 + S_Supplier_3 + S_Supplier_4 + S_Supplier_5
 - rebate_Supplier_1 - rebate_Supplier_2 - rebate_Supplier_3
 - rebate_Supplier_4 - rebate_Supplier_5
Subject To
ActiveLink_Supplier_1_1: x_Supplier_1_1 - 1000000000 z_Supplier_1 <= 0
ActiveLink_Supplier_1_2: x_Supplier_1_2 - 1000000000 z_Supplier_1 <= 0
ActiveLink_Supplier_1_3: x_Supplier_1_3 - 1000000000 z_Supplier_1 <= 0
ActiveLink_Supplier_1_4: x_Supplier_1_4 - 1000000000 z_Supplier_1 <= 0
ActiveLink_Supplier_1_5: x_Supplier_1_5 - 1000000000 z_Supplier_1 <= 0
ActiveLink_Supplier_2_1: x_Supplier_2_1 - 1000000000 z_Supplier_2 <= 0
ActiveLink_Supplier_2_2: x_Supplier_2_2 - 1000000000 z_Supplier_2 <= 0
ActiveLink_Supplier_2_3: x_Supplier_2_3 - 1000000000 z_Supplier_2 <= 0
ActiveLink_Supplier_2_4: x_Supplier_2_4 - 1000000000 z_Supplier_2 <= 0
ActiveLink_Supplier_2_5: x_Supplier_2_5 - 1000000000 z_Supplier_2 <= 0
ActiveLink_Supplier_3_1: x_Supplier_3_1 - 1000000000 z_Supplier_3 <= 0
ActiveLink_Supplier_3_2: x_Supplier_3_2 - 1000000000 z_Supplier_3 <= 0
ActiveLink_Supplier_3_3: x_Supplier_3_3 - 1000000000 z_Supplier_3 <= 0
ActiveLink_Supplier_3_4: x_Supplier_3_4 - 1000000000 z_Supplier_3 <= 0
ActiveLink_Supplier_3_5: x_Supplier_3_5 - 1000000000 z_Supplier_3 <= 0
ActiveLink_Supplier_4_1: x_Supplier_4_1 - 1000000000 z_Supplier_4 <= 0
ActiveLink_Supplier_4_2: x_Supplier_4_2 - 1000000000 z_Supplier_4 <= 0
ActiveLink_Supplier_4_3: x_Supplier_4_3 - 1000000000 z_Supplier_4 <= 0
ActiveLink_Supplier_4_4: x_Supplier_4_4 - 1000000000 z_Supplier_4 <= 0
ActiveLink_Supplier_4_5: x_Supplier_4_5 - 1000000000 z_Supplier_4 <= 0
ActiveLink_Supplier_5_1: x_Supplier_5_1 - 1000000000 z_Supplier_5 <= 0
ActiveLink_Supplier_5_2: x_Supplier_5_2 - 1000000000 z_Supplier_5 <= 0
ActiveLink_Supplier_5_3: x_Supplier_5_3 - 1000000000 z_Supplier_5 <= 0
ActiveLink_Supplier_5_4: x_Supplier_5_4 - 1000000000 z_Supplier_5 <= 0
ActiveLink_Supplier_5_5: x_Supplier_5_5 - 1000000000 z_Supplier_5 <= 0
BaseSpend_Supplier_1: S0_Supplier_1 - 34 x_Supplier_1_1 - 15 x_Supplier_1_2
 - 15 x_Supplier_1_3 - 75 x_Supplier_1_4 - 24 x_Supplier_1_5 = 0
BaseSpend_Supplier_2: S0_Supplier_2 - 14 x_Supplier_2_1 - 15 x_Supplier_2_2
 - 78 x_Supplier_2_3 - 34 x_Supplier_2_4 - 15 x_Supplier_2_5 = 0
BaseSpend_Supplier_3: S0_Supplier_3 - 75 x_Supplier_3_1 - 25 x_Supplier_3_2
 - 56 x_Supplier_3_3 - 24 x_Supplier_3_4 - 87 x_Supplier_3_5 = 0
BaseSpend_Supplier_4: S0_Supplier_4 - 93 x_Supplier_4_1 - 24 x_Supplier_4_2
 - 78 x_Supplier_4_3 - 56 x_Supplier_4_4 - 86 x_Supplier_4_5 = 0
BaseSpend_Supplier_5: S0_Supplier_5 - 76 x_Supplier_5_1 - 32 x_Supplier_5_2
 - 89 x_Supplier_5_3 - 13 x_Supplier_5_4 - 68 x_Supplier_5_5 = 0
Capacity_Supplier_1_Bid_ID_1: x_Supplier_1_1 <= 1457
Capacity_Supplier_1_Bid_ID_2: x_Supplier_1_2 <= 2422
Capacity_Supplier_1_Bid_ID_3: x_Supplier_1_3 <= 6463
Capacity_Supplier_1_Bid_ID_4: x_Supplier_1_4 <= 2358
Capacity_Supplier_1_Bid_ID_5: x_Supplier_1_5 <= 2422
Capacity_Supplier_2_Capacity_Group_Category_1: x_Supplier_2_1 + x_Supplier_2_3
 + x_Supplier_2_4 <= 8001
Capacity_Supplier_2_Capacity_Group_Category_2: x_Supplier_2_2 <= 2501
Capacity_Supplier_2_Capacity_Group_Category_3: x_Supplier_2_5 <= 501
Capacity_Supplier_3_Bid_ID_1: x_Supplier_3_1 <= 1457
Capacity_Supplier_3_Bid_ID_2: x_Supplier_3_2 <= 2422
Capacity_Supplier_3_Bid_ID_3: x_Supplier_3_3 <= 6463
Capacity_Supplier_3_Bid_ID_4: x_Supplier_3_4 <= 2358
Capacity_Supplier_3_Bid_ID_5: x_Supplier_3_5 <= 2422
Capacity_Supplier_4_Bid_ID_1: x_Supplier_4_1 <= 1457
Capacity_Supplier_4_Bid_ID_2: x_Supplier_4_2 <= 2422
Capacity_Supplier_4_Bid_ID_3: x_Supplier_4_3 <= 6463
Capacity_Supplier_4_Bid_ID_4: x_Supplier_4_4 <= 2358
Capacity_Supplier_4_Bid_ID_5: x_Supplier_4_5 <= 2422
Capacity_Supplier_5_Description_Large_Item: x_Supplier_5_5 <= 5001
Capacity_Supplier_5_Description_Medium_item: x_Supplier_5_2 <= 2501
Capacity_Supplier_5_Description_Small_item: x_Supplier_5_1 + x_Supplier_5_3
 + x_Supplier_5_4 <= 9001
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
EffectiveSpend_Supplier_1: - F_Supplier_1 - S0_Supplier_1 + S_Supplier_1
 + d_Supplier_1 = 0
EffectiveSpend_Supplier_2: - F_Supplier_2 - S0_Supplier_2 + S_Supplier_2
 + d_Supplier_2 = 0
EffectiveSpend_Supplier_3: - F_Supplier_3 - S0_Supplier_3 + S_Supplier_3
 + d_Supplier_3 = 0
EffectiveSpend_Supplier_4: - F_Supplier_4 - S0_Supplier_4 + S_Supplier_4
 + d_Supplier_4 = 0
EffectiveSpend_Supplier_5: - F_Supplier_5 - S0_Supplier_5 + S_Supplier_5
 + d_Supplier_5 = 0
Fix_d_Supplier_1: d_Supplier_1 = 0
Fix_d_Supplier_2: d_Supplier_2 = 0
Fix_d_Supplier_3: d_Supplier_3 = 0
Fix_d_Supplier_4: d_Supplier_4 = 0
Fix_d_Supplier_5: d_Supplier_5 = 0
Fix_rebate_Supplier_1: rebate_Supplier_1 = 0
Fix_rebate_Supplier_2: rebate_Supplier_2 = 0
Fix_rebate_Supplier_3: rebate_Supplier_3 = 0
Fix_rebate_Supplier_4: rebate_Supplier_4 = 0
Fix_rebate_Supplier_5: rebate_Supplier_5 = 0
Freight_Supplier_1: F_Supplier_1 - 21 x_Supplier_1_1 - 22 x_Supplier_1_2
 - 22 x_Supplier_1_3 - 22 x_Supplier_1_4 - 20 x_Supplier_1_5 = 0
Freight_Supplier_2: F_Supplier_2 - 18 x_Supplier_2_1 - 14 x_Supplier_2_2
 - 19 x_Supplier_2_3 - 19 x_Supplier_2_4 - 14 x_Supplier_2_5 = 0
Freight_Supplier_3: F_Supplier_3 - 23 x_Supplier_3_1 - 25 x_Supplier_3_2
 - 27 x_Supplier_3_3 - 27 x_Supplier_3_4 - 27 x_Supplier_3_5 = 0
Freight_Supplier_4: F_Supplier_4 - 20 x_Supplier_4_1 - 15 x_Supplier_4_2
 - 18 x_Supplier_4_3 - 20 x_Supplier_4_4 - 19 x_Supplier_4_5 = 0
Freight_Supplier_5: F_Supplier_5 - 30 x_Supplier_5_1 - 32 x_Supplier_5_2
 - 32 x_Supplier_5_3 - 32 x_Supplier_5_4 - 27 x_Supplier_5_5 = 0
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
TransitionLower_1_Supplier_1: x_Supplier_1_1 >= 0
TransitionLower_1_Supplier_3: x_Supplier_3_1 >= 0
TransitionLower_1_Supplier_4: x_Supplier_4_1 >= 0
TransitionLower_1_Supplier_5: x_Supplier_5_1 >= 0
TransitionLower_2_Supplier_1: x_Supplier_1_2 >= 0
TransitionLower_2_Supplier_2: x_Supplier_2_2 >= 0
TransitionLower_2_Supplier_4: x_Supplier_4_2 >= 0
TransitionLower_2_Supplier_5: x_Supplier_5_2 >= 0
TransitionLower_3_Supplier_1: x_Supplier_1_3 >= 0
TransitionLower_3_Supplier_3: x_Supplier_3_3 >= 0
TransitionLower_3_Supplier_4: x_Supplier_4_3 >= 0
TransitionLower_3_Supplier_5: x_Supplier_5_3 >= 0
TransitionLower_4_Supplier_2: x_Supplier_2_4 >= 0
TransitionLower_4_Supplier_3: x_Supplier_3_4 >= 0
TransitionLower_4_Supplier_4: x_Supplier_4_4 >= 0
TransitionLower_4_Supplier_5: x_Supplier_5_4 >= 0
TransitionLower_5_Supplier_1: x_Supplier_1_5 >= 0
TransitionLower_5_Supplier_2: x_Supplier_2_5 >= 0
TransitionLower_5_Supplier_4: x_Supplier_4_5 >= 0
TransitionLower_5_Supplier_5: x_Supplier_5_5 >= 0
Transition_1_Supplier_1: - 1456 T_1_Supplier_1 + x_Supplier_1_1 <= 0
Transition_1_Supplier_3: - 1456 T_1_Supplier_3 + x_Supplier_3_1 <= 0
Transition_1_Supplier_4: - 1456 T_1_Supplier_4 + x_Supplier_4_1 <= 0
Transition_1_Supplier_5: - 1456 T_1_Supplier_5 + x_Supplier_5_1 <= 0
Transition_2_Supplier_1: - 2421 T_2_Supplier_1 + x_Supplier_1_2 <= 0
Transition_2_Supplier_2: - 2421 T_2_Supplier_2 + x_Supplier_2_2 <= 0
Transition_2_Supplier_4: - 2421 T_2_Supplier_4 + x_Supplier_4_2 <= 0
Transition_2_Supplier_5: - 2421 T_2_Supplier_5 + x_Supplier_5_2 <= 0
Transition_3_Supplier_1: - 6462 T_3_Supplier_1 + x_Supplier_1_3 <= 0
Transition_3_Supplier_3: - 6462 T_3_Supplier_3 + x_Supplier_3_3 <= 0
Transition_3_Supplier_4: - 6462 T_3_Supplier_4 + x_Supplier_4_3 <= 0
Transition_3_Supplier_5: - 6462 T_3_Supplier_5 + x_Supplier_5_3 <= 0
Transition_4_Supplier_2: - 2357 T_4_Supplier_2 + x_Supplier_2_4 <= 0
Transition_4_Supplier_3: - 2357 T_4_Supplier_3 + x_Supplier_3_4 <= 0
Transition_4_Supplier_4: - 2357 T_4_Supplier_4 + x_Supplier_4_4 <= 0
Transition_4_Supplier_5: - 2357 T_4_Supplier_5 + x_Supplier_5_4 <= 0
Transition_5_Supplier_1: - 2421 T_5_Supplier_1 + x_Supplier_1_5 <= 0
Transition_5_Supplier_2: - 2421 T_5_Supplier_2 + x_Supplier_2_5 <= 0
Transition_5_Supplier_4: - 2421 T_5_Supplier_4 + x_Supplier_4_5 <= 0
Transition_5_Supplier_5: - 2421 T_5_Supplier_5 + x_Supplier_5_5 <= 0
Volume_Supplier_1: V_Supplier_1 - x_Supplier_1_1 - x_Supplier_1_2
 - x_Supplier_1_3 - x_Supplier_1_4 - x_Supplier_1_5 = 0
Volume_Supplier_2: V_Supplier_2 - x_Supplier_2_1 - x_Supplier_2_2
 - x_Supplier_2_3 - x_Supplier_2_4 - x_Supplier_2_5 = 0
Volume_Supplier_3: V_Supplier_3 - x_Supplier_3_1 - x_Supplier_3_2
 - x_Supplier_3_3 - x_Supplier_3_4 - x_Supplier_3_5 = 0
Volume_Supplier_4: V_Supplier_4 - x_Supplier_4_1 - x_Supplier_4_2
 - x_Supplier_4_3 - x_Supplier_4_4 - x_Supplier_4_5 = 0
Volume_Supplier_5: V_Supplier_5 - x_Supplier_5_1 - x_Supplier_5_2
 - x_Supplier_5_3 - x_Supplier_5_4 - x_Supplier_5_5 = 0
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
z_Supplier_1
z_Supplier_2
z_Supplier_3
z_Supplier_4
z_Supplier_5
End
