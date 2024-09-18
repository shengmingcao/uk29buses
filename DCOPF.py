
from pyomo.environ import *

def DCOPFmodel():

    model=AbstractModel()

    #sets of model
    model.Buses     = Set()  # set of buses
    model.Generator = Set()  # set of generators
    model.Demand    = Set()  # set of demands
    model.Line      = Set()  # set of lines
    model.LE        = Set()  # line-to and from ends set (1,2)
    model.Time      = Set()  # time periods


#Sets
    #Sets of mapping
    model.GenBus = Set(within=model.Buses * model.Generator) #mapping
    model.DemBus = Set(within=model.Buses * model.Demand)  # mapping



# Parameters

    # mapping lines-buses matrix
    model.A = Param(model.Line * model.LE, within=Any)

    # parameters of demand
    model.DT = Param(model.Time, model.Demand, within=Reals)
    model.GT = Param(model.Time,model.Generator)

    # parameters of line capacity and susceptance
    model.SLmax = Param(model.Line, within=NonNegativeReals)
    model.BL = Param(model.Line, within=Reals)  # susceptance of a line

    model.Cost = Param(model.Generator, within=NonNegativeReals)



#Variable
    # control variable
    # output of generator (Variable)
    model.pG = Var(model.Time,model.Generator, domain=NonNegativeReals)


    # state variable
    model.deltaL = Var(model.Line,model.Time, domain=Reals)  # angle difference across lines
    model.delta = Var(model.Buses,model.Time, domain=Reals, initialize=0)  # voltage phase angle at bus b, rad
    model.pL = Var(model.Line,model.Time, domain=Reals)  # real power injected at b onto line l, p.u.

#Constraint
    # --- Kirchoff's current law at each bus b ---
    def KCL_def(model, b,t):
        return sum(model.pG[t,g] for g in model.Generator if (b, g) in model.GenBus) \
            == \
            sum(model.DT[t,d] for d in model.Demand if (b,d) in model.DemBus) + \
            sum(model.pL[l,t] for l in model.Line if model.A[l, 1] == b) - \
            sum(model.pL[l,t] for l in model.Line if model.A[l, 2] == b) \

    model.KCL_const = Constraint(model.Buses,model.Time, rule=KCL_def)

    # Kirchoff's voltage law at each line
    def KVL_line_def(model, l, t):
        return model.pL[l,t] == (-model.BL[l]) * model.deltaL[l,t]

    model.KVL_line_const = Constraint(model.Line, model.Time, rule=KVL_line_def)

    # the power limitations on transmission lines
    def Max_Line_Power1_def(model, l,t):
        return model.pL[l,t] <= model.SLmax[l]
    def Max_Line_Power2_def(model, l,t):
        return model.pL[l,t] >= -model.SLmax[l]

    model.Max_Line_Power1 = Constraint(model.Line,model.Time, rule=Max_Line_Power1_def)
    model.Max_Line_Power2 = Constraint(model.Line,model.Time, rule=Max_Line_Power2_def)


    # the constraints of phase angle
    def phase_angle_dfference(model, l, t):
        return model.deltaL[l, t] == model.delta[model.A[l, 1], t] - \
            model.delta[model.A[l, 2], t]

    model.phase_angle_difference = Constraint(model.Line, model.Time, rule=phase_angle_dfference)


    # Gnenrator output power limitations
    def Max_Real_Power(model, t,g):
        return model.pG[t,g] <= model.GT[t,g]
    def Min_Real_Power(model, t,g):
        return model.pG[t,g] >= (model.GT[t,g]/10)

    model.PGmaxC = Constraint(model.Time,model.Generator, rule=Max_Real_Power)
    model.PGminC = Constraint(model.Time,model.Generator, rule=Min_Real_Power)

# Objective function
    # Cost function
    # Cost function
    def objective(model):
        obj = sum(
            model.Cost[g] * model.pG[g,t] for (g, t) in model.Generator * model.Time)


        return obj

        model.OBJ = Objective(rule=objective, sense=minimize)

    return model











