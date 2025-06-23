using JuMP, Gurobi, XLSX, CSV, DataFrames
# Sets
h = collect(1:8760) # hours of the year
i = collect(1:5) # scenarios for VRE production
# Parameters
# Load the Excel file
xlsx_file = XLSX.readxlsx("FR_DATA_stochastic.xlsx")

# Import data from specific sheets and ranges
alpha = xlsx_file["Load"]["H3:H8762"] # inverse demand function parameter
beta = xlsx_file["Load"]["E7"] # inverse demand function parameter
LF_solar = xlsx_file["Solar"]["B2:F8761"] # Load factor for solar production
LF_onwind = xlsx_file["Onwind"]["B2:F8761"] # Load factor for onshore wind production
LF_offwind = xlsx_file["Offwind"]["B2:F8761"] # Load factor for offshore wind production
hydro_ror = xlsx_file["Hydro_ror"]["B2:B8761"]/1000 # value in GW
avail_nuc = xlsx_file["avail_nuc"]["H2:H8761"] # hourly availability of nuclear power plants

# Inverse demand function parameters
b = 1 / beta
a = alpha*b

# EXTERNAL COSTS
gas_cost = 40 # €/MWh, current gas price in Europe
coal_cost = 10 # €/MWh from Pietzcker
biogas_cost = 90 # €/MWh https://www.pleinchamp.com/actualite/le-gouvernement-revalorise-de-12-le-prix-du-biomethane-injecte 
STOCK_BIO = 20e3 # GWh, half of the injection target of 2030 of 40TWh, from https://www.cre.fr/actualites/nos-lettres-dinformation/la-cre-accompagne-l-essor-du-biomethane-en-france.html 
emis_gas = 0.2 #t/MWh PCS
emis_coal = 0.4 #t/MWh PCS
eff_CCGT = 0.6
eff_OCGT = 0.4
eff_coal = 0.4
eff_BIO = 0.6
eff_BESS = 0.9
eff_PHS = 0.8

eff_IRONAIR = 0.5 # efficiency of iron-air storage
eff_PtG = 0.7 # efficiency of power-to-gas conversion (electrolysis)
eff_H2T = 0.5 # efficiency of hydrogen-to-power conversion (turbine)

discount_rate_nuc = 0.05
discount_rate_low = 0.05
discount_rate_high = 0.09

OPEX_nuc = 0.03 # 3% of investment cost
OPEX_COAL = 0.02
OPEX_CCGT = 0.03
OPEX_OCGT = 0.03
OPEX_BIO = 0.03
OPEX_PV = 0.01
OPEX_ONWIND = 0.03
OPEX_OFFWIND = 0.03
OPEX_BESS = 0.01
OPEX_IRONAIR = 0.01
OPEX_H2 = 0.03

lifetime_nuc = 60 # years
lifetime_COAL = 45
lifetime_CCGT = 45
lifetime_OCGT = 45
lifetime_BIO = 45
lifetime_PV = 25
lifetime_ONWIND = 25
lifetime_OFFWIND = 25
lifetime_BESS = 20
lifetime_IRONAIR = 20
lifetime_UHS = 40
lifetime_PtG = 20
lifetime_H2T = 40

# COST PARAMETERS €/kW
T_nuc = 10 # years of construction for nuclear
T_VRE = 2 # years of construction for VRE
r = 0.05 # discount rate during construction
prem_nuc = ((1+r)^T_nuc-1)/(r*T_nuc)
prem_VRE = ((1+r)^T_VRE-1)/(r*T_VRE)
inv_cost_nuc = 10000*prem_nuc
cycling_limit = 25*2 # current flexibility of French nuclear, double it because we count up and downs
inv_cost_COAL = 2000 # Fraunhofer LCOE 2024
inv_cost_CCGT = 1000
inv_cost_OCGT = 500
inv_cost_BIO = 1000


inv_cost_PV = (955+100)*prem_VRE # IRENA 2023 w/ grid cost externality
inv_cost_ONWIND = (1583+100)*prem_VRE # IRENA 2023
inv_cost_OFFWIND = (3183+100)*prem_VRE # IRENA 2023
inv_cost_BESS_E = 200 # IRENA 2023: cost of 300€/kW
inv_cost_BESS_P = 300 # IRENA 2023: cost of 300€/kW = (300+4*200)/4 approx.


#=
# FORECASTED COSTS for 2035
inv_cost_PV = (600+100)*prem_VRE # 2035 RTE 
inv_cost_ONWIND = (1220+100)*prem_VRE # 2035 RTE 
inv_cost_OFFWIND = (3183*0.7+100)*prem_VRE # 30% reduction cost from now, source = RTE BP 2035
inv_cost_BESS_E = 144 # 2035 from Pietzcker
inv_cost_BESS_P = 122 # 2035 from Pietzcker
=#

# additional technological options (not used in the paper)
inv_cost_IRONAIR_E = 10
inv_cost_IRONAIR_P = 2000 # minesota study
inv_cost_UHS = 5
inv_cost_PtG = 1000
inv_cost_H2T = 1200

# annualized costs in k€/GW
annualized_cost_nuc = inv_cost_nuc*1e3*discount_rate_nuc/(1-(1+discount_rate_nuc)^(-lifetime_nuc)) + OPEX_nuc*inv_cost_nuc
annualized_cost_COAL = inv_cost_COAL*1e3*discount_rate_high/(1-(1+discount_rate_high)^(-lifetime_COAL)) + OPEX_COAL*inv_cost_COAL
annualized_cost_CCGT = inv_cost_CCGT*1e3*discount_rate_high/(1-(1+discount_rate_high)^(-lifetime_CCGT)) + OPEX_CCGT*inv_cost_CCGT
annualized_cost_OCGT = inv_cost_OCGT*1e3*discount_rate_high/(1-(1+discount_rate_high)^(-lifetime_OCGT)) + OPEX_OCGT*inv_cost_OCGT
annualized_cost_BIO = inv_cost_BIO*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_BIO)) + OPEX_BIO*inv_cost_BIO
annualized_cost_PV = inv_cost_PV*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_PV)) + OPEX_PV*inv_cost_PV
annualized_cost_ONWIND = inv_cost_ONWIND*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_ONWIND)) + OPEX_ONWIND*inv_cost_ONWIND
annualized_cost_OFFWIND = inv_cost_OFFWIND*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_OFFWIND)) + OPEX_OFFWIND*inv_cost_OFFWIND
annualized_cost_BESS_E = inv_cost_BESS_E*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_BESS)) + OPEX_BESS*inv_cost_BESS_E
annualized_cost_BESS_P = inv_cost_BESS_P*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_BESS)) + OPEX_BESS*inv_cost_BESS_P
annualized_cost_IRONAIR_E = inv_cost_IRONAIR_E*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_IRONAIR)) + OPEX_IRONAIR*inv_cost_IRONAIR_E
annualized_cost_IRONAIR_P = inv_cost_IRONAIR_P*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_IRONAIR)) + OPEX_IRONAIR*inv_cost_IRONAIR_P
annualized_cost_UHS = inv_cost_UHS*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_UHS)) + OPEX_H2*inv_cost_UHS
annualized_cost_PtG = inv_cost_PtG*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_PtG)) + OPEX_H2*inv_cost_PtG
annualized_cost_H2T = inv_cost_H2T*1e3*discount_rate_low/(1-(1+discount_rate_low)^(-lifetime_H2T)) + OPEX_H2*inv_cost_H2T


CAPA_PHS_P = 8*0.54 # derating factor from Manuel Villavicencio's thesis
CAPA_PHS_E = 8*CAPA_PHS_P*0.54



