using JuMP, Gurobi, XLSX, CSV, DataFrames

# Import data
include("data_loader_stochastic.jl")

function run_model(CO2_cost)
    # costs
    var_cost_CCGT = (gas_cost + CO2_cost*emis_gas)/eff_CCGT
    var_cost_OCGT = (gas_cost + CO2_cost*emis_gas)/eff_OCGT
    var_cost_COAL = (coal_cost + CO2_cost*emis_coal)/eff_coal
    var_cost_BIO = biogas_cost/eff_BIO
    var_cost_NUC = 10
    var_cost_storage = 1 # €/MWh

    # Create model
    model = Model(Gurobi.Optimizer) # if no Gurobi licence, one can use OSQP.Optimizer

    # installed capacity
    @variable(model, CAPA_nuc >= 0)
    @variable(model, CAPA_COAL >= 0)
    @variable(model, CAPA_CCGT >= 0)
    @variable(model, CAPA_OCGT >= 0)
    @variable(model, CAPA_BIO >= 0)
    @variable(model, CAPA_PV >= 0)
    @variable(model, CAPA_onwind >= 0)
    @variable(model, CAPA_offwind >= 0)
    @variable(model, CAPA_BESS_E >= 0)
    @variable(model, CAPA_BESS_P >= 0)
    @variable(model, CAPA_IRONAIR_E >= 0)
    @variable(model, CAPA_IRONAIR_P >= 0)
    @variable(model, CAPA_UHS >= 0)
    @variable(model, CAPA_PtG >= 0)
    @variable(model, CAPA_H2T >= 0)

    # Production variables
    @variable(model, q_nuc[h,i] >= 0)
    @variable(model, q_coal[h,i] >= 0)
    @variable(model, nuc_plus[h,i] >= 0)
    @variable(model, nuc_minus[h,i] >= 0)
    @variable(model, q_hydro[h,i] >= 0)
    @variable(model, q_CCGT[h,i] >= 0)
    @variable(model, q_OCGT[h,i] >= 0)
    @variable(model, q_BIO[h,i] >= 0)
    @variable(model, q_PV[h,i] >= 0)
    @variable(model, q_onwind[h,i] >= 0)
    @variable(model, q_offwind[h,i] >= 0)
    # Demand
    @variable(model, dem[h,i] >= 0)
    # Price & Welfare variables
    @variable(model, lambda[h,i])
    # Storage variables
    @variable(model, charge_BESS[h,i] >= 0)
    @variable(model, discharge_BESS[h,i] >= 0)
    @variable(model, stock_BESS[h,i] >= 0)
    @variable(model, charge_IRONAIR[h,i] >= 0)
    @variable(model, discharge_IRONAIR[h,i] >= 0)
    @variable(model, stock_IRONAIR[h,i] >= 0)
    @variable(model, PtG[h,i] >= 0)
    @variable(model, q_H2T[h,i] >= 0)
    @variable(model, stock_UHS[h,i] >= 0)
    @variable(model, UHS_plus[h,i] >= 0)
    @variable(model, UHS_minus[h,i] >= 0)
    @variable(model, charge_PHS[h,i] >= 0)
    @variable(model, discharge_PHS[h,i] >= 0)
    @variable(model, stock_PHS[h,i] >= 0)

    @variable(model, total_cost >= 0)
    @variable(model, VRE_share >= 0)

    ### EQUATIONS ###

    # Objective function: maximizing welfare
    @objective(model, Max, (1/length(i))*sum((a[h]-0.5*b*dem[h,i])*dem[h,i] - q_nuc[h,i]*var_cost_NUC - q_CCGT[h,i]*var_cost_CCGT - q_OCGT[h,i]*var_cost_OCGT - q_BIO[h,i]*var_cost_BIO -q_coal[h,i]*var_cost_COAL
        - (charge_BESS[h,i]+discharge_BESS[h,i]+charge_IRONAIR[h,i]+discharge_IRONAIR[h,i]+q_H2T[h,i]+PtG[h,i])*var_cost_storage for h in h, i in i) 
        - annualized_cost_nuc*CAPA_nuc 
        - annualized_cost_COAL*CAPA_COAL
        - annualized_cost_CCGT*CAPA_CCGT 
        - annualized_cost_OCGT*CAPA_OCGT 
        - annualized_cost_BIO*CAPA_BIO
        - annualized_cost_PV*CAPA_PV 
        - annualized_cost_ONWIND*CAPA_onwind 
        - annualized_cost_OFFWIND*CAPA_offwind 
        - annualized_cost_BESS_E*CAPA_BESS_E 
        - annualized_cost_BESS_P*CAPA_BESS_P
        - annualized_cost_IRONAIR_E*CAPA_IRONAIR_E 
        - annualized_cost_IRONAIR_P*CAPA_IRONAIR_P
        - annualized_cost_UHS*CAPA_UHS
        - annualized_cost_PtG*CAPA_PtG
        - annualized_cost_H2T*CAPA_H2T
        ) 

    @constraint(model, total_cost == 
        annualized_cost_nuc*CAPA_nuc + annualized_cost_COAL*CAPA_COAL + annualized_cost_CCGT*CAPA_CCGT + 
        annualized_cost_OCGT*CAPA_OCGT + annualized_cost_BIO*CAPA_BIO + annualized_cost_PV*CAPA_PV + 
        annualized_cost_ONWIND*CAPA_onwind + annualized_cost_OFFWIND*CAPA_offwind + 
        annualized_cost_BESS_E*CAPA_BESS_E + annualized_cost_BESS_P*CAPA_BESS_P +
        annualized_cost_IRONAIR_E*CAPA_IRONAIR_E + annualized_cost_IRONAIR_P*CAPA_IRONAIR_P +
        annualized_cost_UHS*CAPA_UHS + annualized_cost_PtG*CAPA_PtG + 
        annualized_cost_H2T*CAPA_H2T +
        sum(q_nuc[h,i]*var_cost_NUC + q_CCGT[h,i]*var_cost_CCGT + q_OCGT[h,i]*var_cost_OCGT + 
            q_BIO[h,i]*var_cost_BIO + q_coal[h,i]*var_cost_COAL + 
            (charge_BESS[h,i] + discharge_BESS[h,i] + charge_IRONAIR[h,i] + discharge_IRONAIR[h,i] + q_H2T[h,i] + PtG[h,i])*var_cost_storage for h in h, i in i) 
        )
    # Demand
    @constraint(model, [h in h, i in i], lambda[h,i] == a[h] - b * dem[h,i])
    # CAPA constraint
    @constraint(model, [h in h, i in i], q_nuc[h,i] <= CAPA_nuc)
    @constraint(model, [h in h, i in i], q_coal[h,i] <= CAPA_COAL)
    @constraint(model, [h in h, i in i], q_CCGT[h,i] <= CAPA_CCGT)
    @constraint(model, [h in h, i in i], q_OCGT[h,i] <= CAPA_OCGT)
    @constraint(model, [h in h, i in i], q_BIO[h,i] <= CAPA_BIO)
    @constraint(model, [i in i], sum(q_BIO[h,i] for h in h) <= STOCK_BIO*eff_BIO) # max stock of biogas

    @constraint(model, [h in h, i in i], q_PV[h,i] <= CAPA_PV * LF_solar[h,i])
    @constraint(model, [h in h, i in i], q_onwind[h,i] <= CAPA_onwind * LF_onwind[h,i])
    @constraint(model, [h in h, i in i], q_offwind[h,i] <= CAPA_offwind * LF_offwind[h,i])
    @constraint(model, [h in h, i in i], charge_BESS[h,i] <= CAPA_BESS_P)
    @constraint(model, [h in h, i in i], discharge_BESS[h,i] <= CAPA_BESS_P)
    @constraint(model, [h in h, i in i], stock_BESS[h,i] <= CAPA_BESS_E)
    @constraint(model, [h in h, i in i], charge_IRONAIR[h,i] <= CAPA_IRONAIR_P)
    @constraint(model, [h in h, i in i], discharge_IRONAIR[h,i] <= CAPA_IRONAIR_P)
    @constraint(model, [h in h, i in i], stock_IRONAIR[h,i] <= CAPA_IRONAIR_E)
    @constraint(model, [h in h, i in i], PtG[h,i] <= CAPA_PtG)
    @constraint(model, [h in h, i in i], q_H2T[h,i] <= CAPA_H2T)
    @constraint(model, [h in h, i in i], stock_UHS[h,i] <= CAPA_UHS)
    @constraint(model, [h in h, i in i], charge_PHS[h,i] <= CAPA_PHS_P)
    @constraint(model, [h in h, i in i], discharge_PHS[h,i] <= CAPA_PHS_P)
    @constraint(model, [h in h, i in i], stock_PHS[h,i] <= CAPA_PHS_E)

    # HYDRO
    @constraint(model, [h in h, i in i], q_hydro[h,i] == hydro_ror[h])

    # ad hoc constraints
    @constraint(model, CAPA_COAL <= 0) # not considering coal in this model
    @constraint(model, CAPA_IRONAIR_E <= 0) # not considering iron-air storage in this model
    @constraint(model, CAPA_UHS <= 0) # not considering hydrogen storage in this model
    #@constraint(model, [h in h, i in i], lambda[h,i] <= 500) # potential price cap for electricity
    @constraint(model, VRE_share == sum(q_PV[h,i] + q_onwind[h,i] + q_offwind[h,i] + q_hydro[h,i] + q_BIO[h,i] for h in h, i in i) / (450e3*5)) # to compute the renewable share in the system

    # Storage
    @constraint(model, [h in 1:length(h)-1, i in i], stock_BESS[h+1,i] - stock_BESS[h,i] - eff_BESS * charge_BESS[h,i] + discharge_BESS[h,i] == 0) 
    @constraint(model, [h in 1:length(h)-1, i in i], stock_IRONAIR[h+1,i] - stock_IRONAIR[h,i] - eff_IRONAIR * charge_IRONAIR[h,i] + discharge_IRONAIR[h,i] == 0) 
    @constraint(model, [h in 1:length(h)-1, i in i], stock_UHS[h+1,i] - stock_UHS[h,i] - eff_PtG * PtG[h,i] + q_H2T[h,i]/eff_H2T == 0)
    @constraint(model, [h in 1:length(h)-1, i in i], stock_PHS[h+1,i] - stock_PHS[h,i] - eff_PHS * charge_PHS[h,i] + discharge_PHS[h,i] == 0) 

    # cycling constraints
    @constraint(model, [h in 1:length(h)-1, i in i], stock_UHS[h+1,i] - stock_UHS[h,i] == UHS_plus[h,i] - UHS_minus[h,i]) # cycling UHS
    @constraint(model, [i in i], sum(UHS_plus[h,i] + UHS_minus[h,i] for h in 1:length(h)-1) <= 10*CAPA_UHS) # cycling UHS

    @constraint(model, [h in 1:length(h)-1, i in i], q_nuc[h+1, i] - q_nuc[h,i] == nuc_plus[h,i] - nuc_minus[h,i]) # cycling nuc
    @constraint(model, [i in i], sum(nuc_plus[h,i] + nuc_minus[h,i] for h in 1:length(h)-1) <= cycling_limit*0.8*CAPA_nuc) # cycling nuc
    @constraint(model, [h in h, i in i], q_nuc[h,i] >= 0.2*CAPA_nuc) # cycling nuc
    @constraint(model, [h in h, i in i], q_nuc[h,i] <= avail_nuc[h]*CAPA_nuc) # cycling nuc

    @constraint(model, [i in i], stock_BESS[1,i] == 0.5*CAPA_BESS_E) # Initial stock of BESS
    @constraint(model, [i in i], stock_PHS[1,i] == 0.5*CAPA_PHS_E) # Initial stock of PHS
    @constraint(model, [i in i], stock_IRONAIR[1,i] == 0.5*CAPA_IRONAIR_E) # Initial stock of IRONAIR
    @constraint(model, [i in i], stock_UHS[1,i] == 0.5*CAPA_UHS) # Initial stock of H2
    @constraint(model, [i in i], stock_BESS[length(h),i] == 0.5*CAPA_BESS_E) # Initial stock of BESS
    @constraint(model, [i in i], stock_PHS[length(h),i] == 0.5*CAPA_PHS_E) # Initial stock of PHS
    @constraint(model, [i in i], stock_IRONAIR[length(h),i] == 0.5*CAPA_IRONAIR_E) # Initial stock of IRONAIR
    @constraint(model, [i in i], stock_UHS[length(h),i] == 0.5*CAPA_UHS) # Initial stock of H2

    # Market clearing
    @constraint(model, [h in h, i in i], 
        dem[h,i] - q_nuc[h,i] -q_coal[h,i] - q_CCGT[h,i] - q_OCGT[h,i] -q_BIO[h,i] - q_PV[h,i] - q_onwind[h,i] - q_offwind[h,i] - q_hydro[h,i] - discharge_BESS[h,i] + charge_BESS[h,i] - discharge_IRONAIR[h,i] + charge_IRONAIR[h,i] - q_H2T[h,i] + PtG[h,i] - discharge_PHS[h,i] + charge_PHS[h,i] == 0)

    # Solve model
    optimize!(model)

    # Convert JuMP variables to DataFrames (hourly values)
    df_dem = DataFrame(dem = value.(dem[:,1].data))
    df_q_nuc = DataFrame(q_nuc = value.(q_nuc[:,1].data))
    df_q_coal = DataFrame(q_coal = value.(q_coal[:,1].data))
    df_lambda = DataFrame(lambda = value.(lambda[:,1].data))
    df_q_onwind = DataFrame(q_onwind = value.(q_onwind[:,1].data))
    df_q_offwind = DataFrame(q_offwind = value.(q_offwind[:,1].data))
    df_q_PV = DataFrame(q_PV = value.(q_PV[:,1].data))

    df_charge_BESS = DataFrame(charge_BESS = value.(charge_BESS[:,1].data))
    df_discharge_BESS = DataFrame(discharge_BESS = value.(discharge_BESS[:,1].data))
    df_stock_BESS = DataFrame(stock_BESS = value.(stock_BESS[:,1].data))

    df_charge_IRONAIR = DataFrame(charge_IRONAIR = value.(charge_IRONAIR[:,1].data))
    df_discharge_IRONAIR = DataFrame(discharge_IRONAIR = value.(discharge_IRONAIR[:,1].data))
    df_stock_IRONAIR = DataFrame(stock_IRONAIR = value.(stock_IRONAIR[:,1].data))

    df_PtG = DataFrame(PtG = value.(PtG[:,1].data))
    df_q_H2T = DataFrame(q_H2T = value.(q_H2T[:,1].data))
    df_stock_UHS = DataFrame(stock_UHS = value.(stock_UHS[:,1].data))

    df_charge_PHS = DataFrame(charge_PHS = value.(charge_PHS[:,1].data))
    df_discharge_PHS = DataFrame(discharge_PHS = value.(discharge_PHS[:,1].data))
    df_stock_PHS = DataFrame(stock_PHS = value.(stock_PHS[:,1].data))

    df_CCGT = DataFrame(q_CCGT = value.(q_CCGT[:,1].data))
    df_OCGT = DataFrame(q_OCGT = value.(q_OCGT[:,1].data))
    df_BIO = DataFrame(q_BIO = value.(q_BIO[:,1].data))
    df_hydro = DataFrame(q_hydro = value.(q_hydro[:,1].data))

    df_total_cost = DataFrame(total_cost = value.(total_cost))
    df_VRE_share = DataFrame(VRE_share = value.(VRE_share))

    # Create a DataFrame for investment capacities
    df_investments = DataFrame(
        Technology = ["Nuclear", "Coal", "CCGT", "OCGT", "BIO", "PV", "Onshore Wind", "Offshore Wind", "BESS Energy", "BESS Power", "IRONAIR Energy", "IRONAIR Power", "UHS", "H2 Turbines", "PtG"],
        Capacity_GW = value.([CAPA_nuc, CAPA_COAL, CAPA_CCGT, CAPA_OCGT, CAPA_BIO, CAPA_PV, CAPA_onwind, CAPA_offwind, CAPA_BESS_E, CAPA_BESS_P, CAPA_IRONAIR_E, CAPA_IRONAIR_P, CAPA_UHS, CAPA_H2T, CAPA_PtG]),
    )
    print(df_investments)

    # Save DataFrames to Excel
    XLSX.openxlsx("RESULTS_stochastic_$CO2_cost.xlsx", mode="w") do xf
        # Add a single sheet for investment decisions
        XLSX.addsheet!(xf, "Investments")
        XLSX.writetable!(xf["Investments"], df_investments)
        # add sheet for total cost
        XLSX.addsheet!(xf, "Total Cost")
        XLSX.writetable!(xf["Total Cost"], df_total_cost)
        # VRE share
        XLSX.addsheet!(xf, "VRE Share")
        XLSX.writetable!(xf["VRE Share"], df_VRE_share)
        # Save time-series data in separate sheets
        for (name, df) in [
            ("dem", df_dem), ("q_nuc", df_q_nuc), ("q_coal", df_q_coal), ("q_CCGT", df_CCGT), ("q_OCGT", df_OCGT), ("q_BIO", df_BIO), 
            ("q_hydro", df_hydro), ("lambda", df_lambda), ("q_onwind", df_q_onwind), 
            ("q_offwind", df_q_offwind), ("q_PV", df_q_PV), ("charge_BESS", df_charge_BESS), 
            ("discharge_BESS", df_discharge_BESS), ("stock_BESS", df_stock_BESS), ("charge_IRONAIR", df_charge_IRONAIR), 
            ("discharge_IRONAIR", df_discharge_IRONAIR), ("stock_IRONAIR", df_stock_IRONAIR), 
            ("PtG", df_PtG), ("q_H2T", df_q_H2T), ("stock_UHS", df_stock_UHS),
            ("charge_PHS", df_charge_PHS), ("discharge_PHS", df_discharge_PHS), 
            ("stock_PHS", df_stock_PHS)
        ]
            XLSX.addsheet!(xf, name)
            XLSX.writetable!(xf[name], df)
        end
    end
end

run_model(0)
#run_model(70)
#run_model(200)
#run_model(500)
#run_model(800)





