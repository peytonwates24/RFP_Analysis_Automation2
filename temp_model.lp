\* Sourcing_with_MultiTier_Rebates_Discounts *\
Minimize
OBJ: S_A + S_B + S_C - rebate_A - rebate_B - rebate_C
Subject To
BaseSpend_A: S0_A - 0.5252 x_A_1 - 64 x_A_10 - 70 x_A_2 - 55 x_A_3 - 23 x_A_4
 - 54 x_A_5 - 42 x_A_6 - 23 x_A_7 - 75 x_A_8 - 97 x_A_9 = 0
BaseSpend_B: S0_B - 10 x_B_1 - 13 x_B_10 - 70 x_B_2 - 65 x_B_3 - 75 x_B_4
 - 34 x_B_5 - 24 x_B_6 - 85 x_B_7 - 13 x_B_8 - 56 x_B_9 = 0
BaseSpend_C: S0_C - 24 x_C_1 - 15 x_C_10 - 75 x_C_2 - 60 x_C_3 - 24 x_C_4
 - 44 x_C_6 - 42 x_C_7 - 24 x_C_8 - 32 x_C_9 = 0
Capacity_B_Bid_ID_1: x_B_1 <= 100000000
Capacity_B_Bid_ID_10: x_B_10 <= 100000000
Capacity_B_Bid_ID_2: x_B_2 <= 100000000
Capacity_B_Bid_ID_3: x_B_3 <= 100000000
Capacity_B_Bid_ID_4: x_B_4 <= 100000000
Capacity_B_Bid_ID_5: x_B_5 <= 100000000
Capacity_B_Bid_ID_6: x_B_6 <= 100000000
Capacity_B_Bid_ID_7: x_B_7 <= 100000000
Capacity_B_Bid_ID_8: x_B_8 <= 100000000
Capacity_B_Bid_ID_9: x_B_9 <= 100000000
Capacity_C_Bid_ID_1: x_C_1 <= 100000000
Capacity_C_Bid_ID_10: x_C_10 <= 100000000
Capacity_C_Bid_ID_2: x_C_2 <= 100000000
Capacity_C_Bid_ID_3: x_C_3 <= 100000000
Capacity_C_Bid_ID_4: x_C_4 <= 100000000
Capacity_C_Bid_ID_5: x_C_5 <= 100000000
Capacity_C_Bid_ID_6: x_C_6 <= 100000000
Capacity_C_Bid_ID_7: x_C_7 <= 100000000
Capacity_C_Bid_ID_8: x_C_8 <= 100000000
Capacity_C_Bid_ID_9: x_C_9 <= 100000000
Demand_1: x_A_1 + x_B_1 + x_C_1 = 700
Demand_10: x_A_10 + x_B_10 + x_C_10 = 13
Demand_2: x_A_2 + x_B_2 + x_C_2 = 9000
Demand_3: x_A_3 + x_B_3 + x_C_3 = 600
Demand_4: x_A_4 + x_B_4 + x_C_4 = 5670
Demand_5: x_A_5 + x_B_5 + x_C_5 = 45
Demand_6: x_A_6 + x_B_6 + x_C_6 = 242
Demand_7: x_A_7 + x_B_7 + x_C_7 = 664
Demand_8: x_A_8 + x_B_8 + x_C_8 = 24
Demand_9: x_A_9 + x_B_9 + x_C_9 = 232
DiscountTierLower_A_0: d_A - 19400000000 z_discount_A_0 >= -19400000000
DiscountTierLower_A_1: - 0.01 S0_A + d_A - 19400000000 z_discount_A_1
 >= -19400000000
DiscountTierLower_B_0: d_B - 97000000000 z_discount_B_0 >= -97000000000
DiscountTierLower_B_1: - 0.03 S0_B + d_B - 97000000000 z_discount_B_1
 >= -97000000000
DiscountTierLower_C_0: d_C - 97000000000 z_discount_C_0 >= -97000000000
DiscountTierLower_C_1: - 0.04 S0_C + d_C - 97000000000 z_discount_C_1
 >= -97000000000
DiscountTierMax_A_0: 19400000000 z_discount_A_0 <= 19400001000
DiscountTierMax_B_0: 97000000000 z_discount_B_0 <= 97000000500
DiscountTierMax_C_0: 97000000000 z_discount_C_0 <= 97000000500
_dummy: __dummy = 0
DiscountTierMin_A_0: __dummy >= 0
DiscountTierMin_A_1: x_A_1 + x_A_10 + x_A_3 + x_A_4 + x_A_8 + x_A_9
 - 1000 z_discount_A_1 >= 0
DiscountTierMin_B_0: __dummy >= 0
DiscountTierMin_B_1: x_B_2 + x_B_5 + x_B_6 + x_B_7 - 500 z_discount_B_1 >= 0
DiscountTierMin_C_0: __dummy >= 0
DiscountTierMin_C_1: x_C_1 + x_C_10 + x_C_3 + x_C_4 + x_C_8 + x_C_9
 - 500 z_discount_C_1 >= 0
DiscountTierSelect_A: z_discount_A_0 + z_discount_A_1 = 1
DiscountTierSelect_B: z_discount_B_0 + z_discount_B_1 = 1
DiscountTierSelect_C: z_discount_C_0 + z_discount_C_1 = 1
DiscountTierUpper_A_0: d_A + 19400000000 z_discount_A_0 <= 19400000000
DiscountTierUpper_A_1: - 0.01 S0_A + d_A + 19400000000 z_discount_A_1
 <= 19400000000
DiscountTierUpper_B_0: d_B + 97000000000 z_discount_B_0 <= 97000000000
DiscountTierUpper_B_1: - 0.03 S0_B + d_B + 97000000000 z_discount_B_1
 <= 97000000000
DiscountTierUpper_C_0: d_C + 97000000000 z_discount_C_0 <= 97000000000
DiscountTierUpper_C_1: - 0.04 S0_C + d_C + 97000000000 z_discount_C_1
 <= 97000000000
EffectiveSpend_A: - S0_A + S_A + d_A = 0
EffectiveSpend_B: - S0_B + S_B + d_B = 0
EffectiveSpend_C: - S0_C + S_C + d_C = 0
NonBid_C_5: x_C_5 = 0
RebateTierLower_A_0: rebate_A - 19400000000 y_rebate_A_0 >= -19400000000
RebateTierLower_A_1: - 0.1 S_A + rebate_A - 19400000000 y_rebate_A_1
 >= -19400000000
RebateTierLower_B_0: rebate_B - 97000000000 y_rebate_B_0 >= -97000000000
RebateTierLower_B_1: - 0.05 S_B + rebate_B - 97000000000 y_rebate_B_1
 >= -97000000000
RebateTierLower_C_0: rebate_C - 97000000000 y_rebate_C_0 >= -97000000000
RebateTierLower_C_1: - 0.07 S_C + rebate_C - 97000000000 y_rebate_C_1
 >= -97000000000
RebateTierMax_A_0: 19400000000 y_rebate_A_0 <= 19400000500
RebateTierMax_B_0: 97000000000 y_rebate_B_0 <= 97000000500
RebateTierMax_C_0: 97000000000 y_rebate_C_0 <= 97000000700
RebateTierMin_A_0: __dummy >= 0
RebateTierMin_A_1: - 500 y_rebate_A_1 >= 0
RebateTierMin_B_0: __dummy >= 0
RebateTierMin_B_1: x_B_2 + x_B_5 + x_B_6 + x_B_7 - 500 y_rebate_B_1 >= 0
RebateTierMin_C_0: __dummy >= 0
RebateTierMin_C_1: x_C_1 + x_C_10 + x_C_3 + x_C_4 + x_C_8 + x_C_9
 - 700 y_rebate_C_1 >= 0
RebateTierSelect_A: y_rebate_A_0 + y_rebate_A_1 = 1
RebateTierSelect_B: y_rebate_B_0 + y_rebate_B_1 = 1
RebateTierSelect_C: y_rebate_C_0 + y_rebate_C_1 = 1
RebateTierUpper_A_0: rebate_A + 19400000000 y_rebate_A_0 <= 19400000000
RebateTierUpper_A_1: - 0.1 S_A + rebate_A + 19400000000 y_rebate_A_1
 <= 19400000000
RebateTierUpper_B_0: rebate_B + 97000000000 y_rebate_B_0 <= 97000000000
RebateTierUpper_B_1: - 0.05 S_B + rebate_B + 97000000000 y_rebate_B_1
 <= 97000000000
RebateTierUpper_C_0: rebate_C + 97000000000 y_rebate_C_0 <= 97000000000
RebateTierUpper_C_1: - 0.07 S_C + rebate_C + 97000000000 y_rebate_C_1
 <= 97000000000
Rule_0_1: 0 x_A_1 - x_B_1 - x_C_1 >= 0
Rule_0_10: - x_A_10 - x_B_10 + 0 x_C_10 >= 0
Rule_0_2: - x_A_2 + 0 x_B_2 - x_C_2 >= 0
Rule_0_3: - x_A_3 - x_B_3 + 0 x_C_3 >= 0
Rule_0_4: - x_A_4 - x_B_4 + 0 x_C_4 >= 0
Rule_0_5: - x_A_5 - x_B_5 + 0 x_C_5 >= 0
Rule_0_6: - x_A_6 - x_B_6 + 0 x_C_6 >= 0
Rule_0_7: - x_A_7 - x_B_7 + 0 x_C_7 >= 0
Rule_0_8: - x_A_8 - x_B_8 + 0 x_C_8 >= 0
Rule_0_9: - x_A_9 - x_B_9 + 0 x_C_9 >= 0
Transition_10_A: - 13 T_10_A + x_A_10 <= 0
Transition_10_B: - 13 T_10_B + x_B_10 <= 0
Transition_1_B: - 700 T_1_B + x_B_1 <= 0
Transition_1_C: - 700 T_1_C + x_C_1 <= 0
Transition_2_A: - 9000 T_2_A + x_A_2 <= 0
Transition_2_C: - 9000 T_2_C + x_C_2 <= 0
Transition_3_A: - 600 T_3_A + x_A_3 <= 0
Transition_3_B: - 600 T_3_B + x_B_3 <= 0
Transition_4_A: - 5670 T_4_A + x_A_4 <= 0
Transition_4_B: - 5670 T_4_B + x_B_4 <= 0
Transition_5_A: - 45 T_5_A + x_A_5 <= 0
Transition_5_B: - 45 T_5_B + x_B_5 <= 0
Transition_6_A: - 242 T_6_A + x_A_6 <= 0
Transition_6_B: - 242 T_6_B + x_B_6 <= 0
Transition_7_A: - 664 T_7_A + x_A_7 <= 0
Transition_7_B: - 664 T_7_B + x_B_7 <= 0
Transition_8_A: - 24 T_8_A + x_A_8 <= 0
Transition_8_B: - 24 T_8_B + x_B_8 <= 0
Transition_9_A: - 232 T_9_A + x_A_9 <= 0
Transition_9_B: - 232 T_9_B + x_B_9 <= 0
Volume_A: V_A - x_A_1 - x_A_10 - x_A_2 - x_A_3 - x_A_4 - x_A_5 - x_A_6 - x_A_7
 - x_A_8 - x_A_9 = 0
Volume_B: V_B - x_B_1 - x_B_10 - x_B_2 - x_B_3 - x_B_4 - x_B_5 - x_B_6 - x_B_7
 - x_B_8 - x_B_9 = 0
Volume_C: V_C - x_C_1 - x_C_10 - x_C_2 - x_C_3 - x_C_4 - x_C_5 - x_C_6 - x_C_7
 - x_C_8 - x_C_9 = 0
Binaries
T_10_A
T_10_B
T_1_B
T_1_C
T_2_A
T_2_C
T_3_A
T_3_B
T_4_A
T_4_B
T_5_A
T_5_B
T_6_A
T_6_B
T_7_A
T_7_B
T_8_A
T_8_B
T_9_A
T_9_B
y_rebate_A_0
y_rebate_A_1
y_rebate_B_0
y_rebate_B_1
y_rebate_C_0
y_rebate_C_1
z_discount_A_0
z_discount_A_1
z_discount_B_0
z_discount_B_1
z_discount_C_0
z_discount_C_1
End
