"""
Reconciliation Optimizer using PuLP for Excel

Authors: A.J. Thompson

In order for reconciler to function, you need to first index your data as follows:

1) Creat one array of unique key values for each G/L transaction and another array for each bank transaction
2) Map these arrays respectively as GL_keys and Bank_keys
3) Create separate arrays for each set of G/L and bank transaction amounts, GL_transactions and Bank_transactions
4) Create separate arrays for each set of G/L and bank dates, GL_dates and Bank_dates
5) Create multi-dimensional array for your binary decision values, transactions_reconciled

"""
# Import PuLP modeler functions
from pulp import *

# The prob variable is created to contain the problem data
prob = LpProblem("Reconciliation Problem", LpMaximize)

# Creates a list of tuples containing all possible reconciled GL items
possible_reconciled = [(g, b) for g in GL_keys for b in Bank_keys]

# A dictionary called 'vars' is created to contain the referenced variables (the possible reconciled)
vars = LpVariable.dicts("Possible Reconciled", (GL_keys, Bank_keys), 0, 1, LpBinary)

# The main objective function is added: maximize the amount reconciled
prob += lpSum([vars[g][b] for (g, b) in possible_reconciled]), "Sum_of_Reconciled_Items"

# Add constraint that will allow each GL transaction to be reconciled only once
for g in GL_keys:
    prob += lpSum([vars[g][b] for b in Bank_keys]) <= 1
    prob += lpSum([vars[g][b] * GL_dates[g] for b in Bank_keys]
                  ) <= ([vars[g][b] * Bank_dates[b] for b in Bank_keys])

# Add constraint to keep sum of GL items from exceeding value of any one bank transaction
for b in Bank_keys:
    prob += lpSum([vars[g][b] * GL_transactions[g] for g in GL_keys]) <= Bank_transactions[b]

# The problem data is written to an .lp file
prob.writeLP("ReconciliationProblem.lp")

# The problem is solved using PuLP's choice of Solver
prob.solve(COIN_CMD(msg=1, maxSeconds=1500))

# The status of the solution is printed to the screen
print "Status:", LpStatus[prob.status]

# The optimised objective function value is printed to the screen
print "Total Reconciled Items = ", value(prob.objective)

# Print solved reconciliation onto Excel transactions_reconciled table
for (g, b) in possible_reconciled:
    transactions_reconciled[g, b] = vars[g][b].varValue

# Add status of problem to SolverResult variable for further use
SolverResult = LpStatus[prob.status]
