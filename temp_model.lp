\* Sourcing_with_MultiTier_Rebates_Discounts *\
Minimize
OBJ: S_Chemtrade + S_SATCO + S_Shrieve + S_Southern_State - rebate_Chemtrade
 - rebate_SATCO - rebate_Shrieve - rebate_Southern_State
Subject To
BaseSpend_Chemtrade: S0_Chemtrade - 196.8 x_Chemtrade_11
 - 343.8 x_Chemtrade_15 - 306.8 x_Chemtrade_8 - 195.8 x_Chemtrade_9 = 0
BaseSpend_SATCO: S0_SATCO - 217.8 x_SATCO_11 - 218.8 x_SATCO_18
 - 159.8 x_SATCO_23 - 205.8 x_SATCO_6 = 0
BaseSpend_Shrieve: S0_Shrieve - 132.8 x_Shrieve_1 - 175.8 x_Shrieve_10
 - 160.8 x_Shrieve_11 - 132.8 x_Shrieve_13 - 135.8 x_Shrieve_16
 - 164.97 x_Shrieve_18 - 184.3 x_Shrieve_23 - 139 x_Shrieve_3
 - 131.81 x_Shrieve_5 - 128.8 x_Shrieve_6 - 171.8 x_Shrieve_9 = 0
BaseSpend_Southern_State: S0_Southern_State - 188.38 x_Southern_State_1
 - 190.88 x_Southern_State_16 - 253.46 x_Southern_State_17
 - 253.46 x_Southern_State_2 - 150.46 x_Southern_State_20
 - 131.8 x_Southern_State_23 - 187.12 x_Southern_State_3
 - 253.46 x_Southern_State_4 - 139.84 x_Southern_State_5
 - 160.49 x_Southern_State_6 = 0
Capacity_Chemtrade_Bid_ID_11: x_Chemtrade_11 <= 6700
Capacity_Chemtrade_Bid_ID_15: x_Chemtrade_15 <= 660
Capacity_Chemtrade_Bid_ID_8: x_Chemtrade_8 <= 480
Capacity_Chemtrade_Bid_ID_9: x_Chemtrade_9 <= 19000
Capacity_SATCO_Bid_ID_11: x_SATCO_11 <= 6700
Capacity_SATCO_Bid_ID_18: x_SATCO_18 <= 13500
Capacity_SATCO_Bid_ID_23: x_SATCO_23 <= 9000
Capacity_SATCO_Bid_ID_6: x_SATCO_6 <= 21200
Capacity_Shrieve_Bid_ID_1: x_Shrieve_1 <= 48000
Capacity_Shrieve_Bid_ID_10: x_Shrieve_10 <= 10500
Capacity_Shrieve_Bid_ID_11: x_Shrieve_11 <= 6700
Capacity_Shrieve_Bid_ID_13: x_Shrieve_13 <= 26400
Capacity_Shrieve_Bid_ID_16: x_Shrieve_16 <= 26000
Capacity_Shrieve_Bid_ID_18: x_Shrieve_18 <= 13500
Capacity_Shrieve_Bid_ID_23: x_Shrieve_23 <= 9000
Capacity_Shrieve_Bid_ID_3: x_Shrieve_3 <= 17000
Capacity_Shrieve_Bid_ID_5: x_Shrieve_5 <= 35000
Capacity_Shrieve_Bid_ID_6: x_Shrieve_6 <= 21200
Capacity_Shrieve_Bid_ID_9: x_Shrieve_9 <= 19000
Capacity_Southern_State_Bid_ID_1: x_Southern_State_1 <= 48000
Capacity_Southern_State_Bid_ID_16: x_Southern_State_16 <= 26000
Capacity_Southern_State_Bid_ID_17: x_Southern_State_17 <= 2000
Capacity_Southern_State_Bid_ID_2: x_Southern_State_2 <= 4000
Capacity_Southern_State_Bid_ID_20: x_Southern_State_20 <= 670
Capacity_Southern_State_Bid_ID_23: x_Southern_State_23 <= 9000
Capacity_Southern_State_Bid_ID_3: x_Southern_State_3 <= 17000
Capacity_Southern_State_Bid_ID_4: x_Southern_State_4 <= 12500
Capacity_Southern_State_Bid_ID_5: x_Southern_State_5 <= 35000
Capacity_Southern_State_Bid_ID_6: x_Southern_State_6 <= 21200
Demand_1: x_Chemtrade_1 + x_SATCO_1 + x_Shrieve_1 + x_Southern_State_1 = 48000
Demand_10: x_Chemtrade_10 + x_SATCO_10 + x_Shrieve_10 + x_Southern_State_10
 = 10500
Demand_11: x_Chemtrade_11 + x_SATCO_11 + x_Shrieve_11 + x_Southern_State_11
 = 6700
Demand_12: x_Chemtrade_12 + x_SATCO_12 + x_Shrieve_12 + x_Southern_State_12
 = 900
Demand_13: x_Chemtrade_13 + x_SATCO_13 + x_Shrieve_13 + x_Southern_State_13
 = 26400
Demand_14: x_Chemtrade_14 + x_SATCO_14 + x_Shrieve_14 + x_Southern_State_14
 = 2300
Demand_15: x_Chemtrade_15 + x_SATCO_15 + x_Shrieve_15 + x_Southern_State_15
 = 660
Demand_16: x_Chemtrade_16 + x_SATCO_16 + x_Shrieve_16 + x_Southern_State_16
 = 26000
Demand_17: x_Chemtrade_17 + x_SATCO_17 + x_Shrieve_17 + x_Southern_State_17
 = 2000
Demand_18: x_Chemtrade_18 + x_SATCO_18 + x_Shrieve_18 + x_Southern_State_18
 = 13500
Demand_19: x_Chemtrade_19 + x_SATCO_19 + x_Shrieve_19 + x_Southern_State_19
 = 300
Demand_2: x_Chemtrade_2 + x_SATCO_2 + x_Shrieve_2 + x_Southern_State_2 = 4000
Demand_20: x_Chemtrade_20 + x_SATCO_20 + x_Shrieve_20 + x_Southern_State_20
 = 670
Demand_21: x_Chemtrade_21 + x_SATCO_21 + x_Shrieve_21 + x_Southern_State_21
 = 8500
Demand_22: x_Chemtrade_22 + x_SATCO_22 + x_Shrieve_22 + x_Southern_State_22
 = 7200
Demand_23: x_Chemtrade_23 + x_SATCO_23 + x_Shrieve_23 + x_Southern_State_23
 = 9000
Demand_3: x_Chemtrade_3 + x_SATCO_3 + x_Shrieve_3 + x_Southern_State_3 = 17000
Demand_4: x_Chemtrade_4 + x_SATCO_4 + x_Shrieve_4 + x_Southern_State_4 = 12500
Demand_5: x_Chemtrade_5 + x_SATCO_5 + x_Shrieve_5 + x_Southern_State_5 = 35000
Demand_6: x_Chemtrade_6 + x_SATCO_6 + x_Shrieve_6 + x_Southern_State_6 = 21200
Demand_7: x_Chemtrade_7 + x_SATCO_7 + x_Shrieve_7 + x_Southern_State_7 = 350
Demand_8: x_Chemtrade_8 + x_SATCO_8 + x_Shrieve_8 + x_Southern_State_8 = 480
Demand_9: x_Chemtrade_9 + x_SATCO_9 + x_Shrieve_9 + x_Southern_State_9 = 19000
EffectiveSpend_Chemtrade: - S0_Chemtrade + S_Chemtrade + d_Chemtrade = 0
EffectiveSpend_SATCO: - S0_SATCO + S_SATCO + d_SATCO = 0
EffectiveSpend_Shrieve: - S0_Shrieve + S_Shrieve + d_Shrieve = 0
EffectiveSpend_Southern_State: - S0_Southern_State + S_Southern_State
 + d_Southern_State = 0
Fix_d_Chemtrade: d_Chemtrade = 0
Fix_d_SATCO: d_SATCO = 0
Fix_d_Shrieve: d_Shrieve = 0
Fix_d_Southern_State: d_Southern_State = 0
Fix_rebate_Chemtrade: rebate_Chemtrade = 0
Fix_rebate_SATCO: rebate_SATCO = 0
Fix_rebate_Shrieve: rebate_Shrieve = 0
Fix_rebate_Southern_State: rebate_Southern_State = 0
NonBid_Chemtrade_1: x_Chemtrade_1 = 0
NonBid_Chemtrade_10: x_Chemtrade_10 = 0
NonBid_Chemtrade_12: x_Chemtrade_12 = 0
NonBid_Chemtrade_13: x_Chemtrade_13 = 0
NonBid_Chemtrade_14: x_Chemtrade_14 = 0
NonBid_Chemtrade_16: x_Chemtrade_16 = 0
NonBid_Chemtrade_17: x_Chemtrade_17 = 0
NonBid_Chemtrade_18: x_Chemtrade_18 = 0
NonBid_Chemtrade_19: x_Chemtrade_19 = 0
NonBid_Chemtrade_2: x_Chemtrade_2 = 0
NonBid_Chemtrade_20: x_Chemtrade_20 = 0
NonBid_Chemtrade_21: x_Chemtrade_21 = 0
NonBid_Chemtrade_22: x_Chemtrade_22 = 0
NonBid_Chemtrade_23: x_Chemtrade_23 = 0
NonBid_Chemtrade_3: x_Chemtrade_3 = 0
NonBid_Chemtrade_4: x_Chemtrade_4 = 0
NonBid_Chemtrade_5: x_Chemtrade_5 = 0
NonBid_Chemtrade_6: x_Chemtrade_6 = 0
NonBid_Chemtrade_7: x_Chemtrade_7 = 0
NonBid_SATCO_1: x_SATCO_1 = 0
NonBid_SATCO_10: x_SATCO_10 = 0
NonBid_SATCO_12: x_SATCO_12 = 0
NonBid_SATCO_13: x_SATCO_13 = 0
NonBid_SATCO_14: x_SATCO_14 = 0
NonBid_SATCO_15: x_SATCO_15 = 0
NonBid_SATCO_16: x_SATCO_16 = 0
NonBid_SATCO_17: x_SATCO_17 = 0
NonBid_SATCO_19: x_SATCO_19 = 0
NonBid_SATCO_2: x_SATCO_2 = 0
NonBid_SATCO_20: x_SATCO_20 = 0
NonBid_SATCO_21: x_SATCO_21 = 0
NonBid_SATCO_22: x_SATCO_22 = 0
NonBid_SATCO_3: x_SATCO_3 = 0
NonBid_SATCO_4: x_SATCO_4 = 0
NonBid_SATCO_5: x_SATCO_5 = 0
NonBid_SATCO_7: x_SATCO_7 = 0
NonBid_SATCO_8: x_SATCO_8 = 0
NonBid_SATCO_9: x_SATCO_9 = 0
NonBid_Shrieve_12: x_Shrieve_12 = 0
NonBid_Shrieve_14: x_Shrieve_14 = 0
NonBid_Shrieve_15: x_Shrieve_15 = 0
NonBid_Shrieve_17: x_Shrieve_17 = 0
NonBid_Shrieve_19: x_Shrieve_19 = 0
NonBid_Shrieve_2: x_Shrieve_2 = 0
NonBid_Shrieve_20: x_Shrieve_20 = 0
NonBid_Shrieve_21: x_Shrieve_21 = 0
NonBid_Shrieve_22: x_Shrieve_22 = 0
NonBid_Shrieve_4: x_Shrieve_4 = 0
NonBid_Shrieve_7: x_Shrieve_7 = 0
NonBid_Shrieve_8: x_Shrieve_8 = 0
NonBid_Southern_State_10: x_Southern_State_10 = 0
NonBid_Southern_State_11: x_Southern_State_11 = 0
NonBid_Southern_State_12: x_Southern_State_12 = 0
NonBid_Southern_State_13: x_Southern_State_13 = 0
NonBid_Southern_State_14: x_Southern_State_14 = 0
NonBid_Southern_State_15: x_Southern_State_15 = 0
NonBid_Southern_State_18: x_Southern_State_18 = 0
NonBid_Southern_State_19: x_Southern_State_19 = 0
NonBid_Southern_State_21: x_Southern_State_21 = 0
NonBid_Southern_State_22: x_Southern_State_22 = 0
NonBid_Southern_State_7: x_Southern_State_7 = 0
NonBid_Southern_State_8: x_Southern_State_8 = 0
NonBid_Southern_State_9: x_Southern_State_9 = 0
Rule_0_1: - 0.5 x_Chemtrade_1 - 0.5 x_SATCO_1 + 0.5 x_Shrieve_1
 - 0.5 x_Southern_State_1 >= 0
Transition_10_Chemtrade: - 10500 T_10_Chemtrade + x_Chemtrade_10 <= 0
Transition_10_SATCO: - 10500 T_10_SATCO + x_SATCO_10 <= 0
Transition_10_Southern_State: - 10500 T_10_Southern_State
 + x_Southern_State_10 <= 0
Transition_11_SATCO: - 6700 T_11_SATCO + x_SATCO_11 <= 0
Transition_11_Shrieve: - 6700 T_11_Shrieve + x_Shrieve_11 <= 0
Transition_11_Southern_State: - 6700 T_11_Southern_State + x_Southern_State_11
 <= 0
Transition_12_Chemtrade: - 900 T_12_Chemtrade + x_Chemtrade_12 <= 0
Transition_12_Shrieve: - 900 T_12_Shrieve + x_Shrieve_12 <= 0
Transition_12_Southern_State: - 900 T_12_Southern_State + x_Southern_State_12
 <= 0
Transition_13_Chemtrade: - 26400 T_13_Chemtrade + x_Chemtrade_13 <= 0
Transition_13_SATCO: - 26400 T_13_SATCO + x_SATCO_13 <= 0
Transition_13_Southern_State: - 26400 T_13_Southern_State
 + x_Southern_State_13 <= 0
Transition_14_Chemtrade: - 2300 T_14_Chemtrade + x_Chemtrade_14 <= 0
Transition_14_Shrieve: - 2300 T_14_Shrieve + x_Shrieve_14 <= 0
Transition_14_Southern_State: - 2300 T_14_Southern_State + x_Southern_State_14
 <= 0
Transition_15_SATCO: - 660 T_15_SATCO + x_SATCO_15 <= 0
Transition_15_Shrieve: - 660 T_15_Shrieve + x_Shrieve_15 <= 0
Transition_15_Southern_State: - 660 T_15_Southern_State + x_Southern_State_15
 <= 0
Transition_16_Chemtrade: - 26000 T_16_Chemtrade + x_Chemtrade_16 <= 0
Transition_16_SATCO: - 26000 T_16_SATCO + x_SATCO_16 <= 0
Transition_16_Southern_State: - 26000 T_16_Southern_State
 + x_Southern_State_16 <= 0
Transition_17_Chemtrade: - 2000 T_17_Chemtrade + x_Chemtrade_17 <= 0
Transition_17_SATCO: - 2000 T_17_SATCO + x_SATCO_17 <= 0
Transition_17_Shrieve: - 2000 T_17_Shrieve + x_Shrieve_17 <= 0
Transition_18_Chemtrade: - 13500 T_18_Chemtrade + x_Chemtrade_18 <= 0
Transition_18_SATCO: - 13500 T_18_SATCO + x_SATCO_18 <= 0
Transition_18_Southern_State: - 13500 T_18_Southern_State
 + x_Southern_State_18 <= 0
Transition_19_Chemtrade: - 300 T_19_Chemtrade + x_Chemtrade_19 <= 0
Transition_19_Shrieve: - 300 T_19_Shrieve + x_Shrieve_19 <= 0
Transition_19_Southern_State: - 300 T_19_Southern_State + x_Southern_State_19
 <= 0
Transition_1_Chemtrade: - 48000 T_1_Chemtrade + x_Chemtrade_1 <= 0
Transition_1_SATCO: - 48000 T_1_SATCO + x_SATCO_1 <= 0
Transition_1_Southern_State: - 48000 T_1_Southern_State + x_Southern_State_1
 <= 0
Transition_20_Chemtrade: - 670 T_20_Chemtrade + x_Chemtrade_20 <= 0
Transition_20_SATCO: - 670 T_20_SATCO + x_SATCO_20 <= 0
Transition_20_Shrieve: - 670 T_20_Shrieve + x_Shrieve_20 <= 0
Transition_21_SATCO: - 8500 T_21_SATCO + x_SATCO_21 <= 0
Transition_21_Shrieve: - 8500 T_21_Shrieve + x_Shrieve_21 <= 0
Transition_21_Southern_State: - 8500 T_21_Southern_State + x_Southern_State_21
 <= 0
Transition_22_SATCO: - 7200 T_22_SATCO + x_SATCO_22 <= 0
Transition_22_Shrieve: - 7200 T_22_Shrieve + x_Shrieve_22 <= 0
Transition_22_Southern_State: - 7200 T_22_Southern_State + x_Southern_State_22
 <= 0
Transition_23_Chemtrade: - 9000 T_23_Chemtrade + x_Chemtrade_23 <= 0
Transition_23_SATCO: - 9000 T_23_SATCO + x_SATCO_23 <= 0
Transition_23_Southern_State: - 9000 T_23_Southern_State + x_Southern_State_23
 <= 0
Transition_2_Chemtrade: - 4000 T_2_Chemtrade + x_Chemtrade_2 <= 0
Transition_2_SATCO: - 4000 T_2_SATCO + x_SATCO_2 <= 0
Transition_2_Shrieve: - 4000 T_2_Shrieve + x_Shrieve_2 <= 0
Transition_3_Chemtrade: - 17000 T_3_Chemtrade + x_Chemtrade_3 <= 0
Transition_3_SATCO: - 17000 T_3_SATCO + x_SATCO_3 <= 0
Transition_3_Southern_State: - 17000 T_3_Southern_State + x_Southern_State_3
 <= 0
Transition_4_Chemtrade: - 12500 T_4_Chemtrade + x_Chemtrade_4 <= 0
Transition_4_SATCO: - 12500 T_4_SATCO + x_SATCO_4 <= 0
Transition_4_Shrieve: - 12500 T_4_Shrieve + x_Shrieve_4 <= 0
Transition_5_Chemtrade: - 35000 T_5_Chemtrade + x_Chemtrade_5 <= 0
Transition_5_SATCO: - 35000 T_5_SATCO + x_SATCO_5 <= 0
Transition_5_Southern_State: - 35000 T_5_Southern_State + x_Southern_State_5
 <= 0
Transition_6_Chemtrade: - 21200 T_6_Chemtrade + x_Chemtrade_6 <= 0
Transition_6_SATCO: - 21200 T_6_SATCO + x_SATCO_6 <= 0
Transition_6_Southern_State: - 21200 T_6_Southern_State + x_Southern_State_6
 <= 0
Transition_7_Chemtrade: - 350 T_7_Chemtrade + x_Chemtrade_7 <= 0
Transition_7_Shrieve: - 350 T_7_Shrieve + x_Shrieve_7 <= 0
Transition_7_Southern_State: - 350 T_7_Southern_State + x_Southern_State_7
 <= 0
Transition_8_SATCO: - 480 T_8_SATCO + x_SATCO_8 <= 0
Transition_8_Shrieve: - 480 T_8_Shrieve + x_Shrieve_8 <= 0
Transition_8_Southern_State: - 480 T_8_Southern_State + x_Southern_State_8
 <= 0
Transition_9_SATCO: - 19000 T_9_SATCO + x_SATCO_9 <= 0
Transition_9_Shrieve: - 19000 T_9_Shrieve + x_Shrieve_9 <= 0
Transition_9_Southern_State: - 19000 T_9_Southern_State + x_Southern_State_9
 <= 0
Volume_Chemtrade: V_Chemtrade - x_Chemtrade_1 - x_Chemtrade_10
 - x_Chemtrade_11 - x_Chemtrade_12 - x_Chemtrade_13 - x_Chemtrade_14
 - x_Chemtrade_15 - x_Chemtrade_16 - x_Chemtrade_17 - x_Chemtrade_18
 - x_Chemtrade_19 - x_Chemtrade_2 - x_Chemtrade_20 - x_Chemtrade_21
 - x_Chemtrade_22 - x_Chemtrade_23 - x_Chemtrade_3 - x_Chemtrade_4
 - x_Chemtrade_5 - x_Chemtrade_6 - x_Chemtrade_7 - x_Chemtrade_8
 - x_Chemtrade_9 = 0
Volume_SATCO: V_SATCO - x_SATCO_1 - x_SATCO_10 - x_SATCO_11 - x_SATCO_12
 - x_SATCO_13 - x_SATCO_14 - x_SATCO_15 - x_SATCO_16 - x_SATCO_17 - x_SATCO_18
 - x_SATCO_19 - x_SATCO_2 - x_SATCO_20 - x_SATCO_21 - x_SATCO_22 - x_SATCO_23
 - x_SATCO_3 - x_SATCO_4 - x_SATCO_5 - x_SATCO_6 - x_SATCO_7 - x_SATCO_8
 - x_SATCO_9 = 0
Volume_Shrieve: V_Shrieve - x_Shrieve_1 - x_Shrieve_10 - x_Shrieve_11
 - x_Shrieve_12 - x_Shrieve_13 - x_Shrieve_14 - x_Shrieve_15 - x_Shrieve_16
 - x_Shrieve_17 - x_Shrieve_18 - x_Shrieve_19 - x_Shrieve_2 - x_Shrieve_20
 - x_Shrieve_21 - x_Shrieve_22 - x_Shrieve_23 - x_Shrieve_3 - x_Shrieve_4
 - x_Shrieve_5 - x_Shrieve_6 - x_Shrieve_7 - x_Shrieve_8 - x_Shrieve_9 = 0
Volume_Southern_State: V_Southern_State - x_Southern_State_1
 - x_Southern_State_10 - x_Southern_State_11 - x_Southern_State_12
 - x_Southern_State_13 - x_Southern_State_14 - x_Southern_State_15
 - x_Southern_State_16 - x_Southern_State_17 - x_Southern_State_18
 - x_Southern_State_19 - x_Southern_State_2 - x_Southern_State_20
 - x_Southern_State_21 - x_Southern_State_22 - x_Southern_State_23
 - x_Southern_State_3 - x_Southern_State_4 - x_Southern_State_5
 - x_Southern_State_6 - x_Southern_State_7 - x_Southern_State_8
 - x_Southern_State_9 = 0
Binaries
T_10_Chemtrade
T_10_SATCO
T_10_Southern_State
T_11_SATCO
T_11_Shrieve
T_11_Southern_State
T_12_Chemtrade
T_12_Shrieve
T_12_Southern_State
T_13_Chemtrade
T_13_SATCO
T_13_Southern_State
T_14_Chemtrade
T_14_Shrieve
T_14_Southern_State
T_15_SATCO
T_15_Shrieve
T_15_Southern_State
T_16_Chemtrade
T_16_SATCO
T_16_Southern_State
T_17_Chemtrade
T_17_SATCO
T_17_Shrieve
T_18_Chemtrade
T_18_SATCO
T_18_Southern_State
T_19_Chemtrade
T_19_Shrieve
T_19_Southern_State
T_1_Chemtrade
T_1_SATCO
T_1_Southern_State
T_20_Chemtrade
T_20_SATCO
T_20_Shrieve
T_21_SATCO
T_21_Shrieve
T_21_Southern_State
T_22_SATCO
T_22_Shrieve
T_22_Southern_State
T_23_Chemtrade
T_23_SATCO
T_23_Southern_State
T_2_Chemtrade
T_2_SATCO
T_2_Shrieve
T_3_Chemtrade
T_3_SATCO
T_3_Southern_State
T_4_Chemtrade
T_4_SATCO
T_4_Shrieve
T_5_Chemtrade
T_5_SATCO
T_5_Southern_State
T_6_Chemtrade
T_6_SATCO
T_6_Southern_State
T_7_Chemtrade
T_7_Shrieve
T_7_Southern_State
T_8_SATCO
T_8_Shrieve
T_8_Southern_State
T_9_SATCO
T_9_Shrieve
T_9_Southern_State
End
