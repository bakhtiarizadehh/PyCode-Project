import win32com.client as com
import numpy as np
import pandas as pd
from openpyxl import load_workbook
import VisumPy.helpers as h

Visum = com.Dispatch("Visum.Visum.250")

versionPath = r'D:\italy\Semester4\Resilience\ResilienceProject\Resilience_Base_Network.ver'
Visum.LoadVersion(versionPath)

LinkList = Visum.Net.Links

# ------------------------------------------------------------
# NETWORK DATA
# ------------------------------------------------------------
zones = Visum.Net.Zones.GetMultiAttValues("NO")
capacity_value = np.array(h.GetMulti(Visum.Net.Links, "CAPPRT"), dtype=float)
from_node = Visum.Net.Links.GetMultiAttValues("FROMNODENO")
to_node = Visum.Net.Links.GetMultiAttValues("TONODENO")
link_type = np.array(h.GetMulti(Visum.Net.Links, "TYPENO"),dtype=int)

original_capacity = np.copy(capacity_value)

# ------------------------------------------------------------
# MONTE CARLO SETTINGS
# ------------------------------------------------------------
n_simulations = 2
np.random.seed(2026)

# High-risk: Beta(4,6), mean = 0.40
# Low-risk:  Beta(2,18), mean = 0.10

high_alpha = 4
high_beta = 6

low_alpha = 2
low_beta = 18

# Capacity remaining when a link fails
failed_capacity = 0.001

# ------------------------------------------------------------
# OUTPUT
# ------------------------------------------------------------

file_name = "resilience_monte_carlo.xlsx"

# ------------------------------------------------------------
# BASELINE
# ------------------------------------------------------------

# Make sure the original network is used
h.SetMulti(Visum.Net.Links,"CAPPRT",original_capacity)
Proc = Visum.Procedures

# Run normal network
Proc.Execute()
flow = np.array(h.GetMulti(Visum.Net.Links, "VolVehPrT(AP)"))
travel_time = np.array(h.GetMulti(Visum.Net.Links, "TCur_PrTSys(C)"))

# Total travel time of normal network
baseline_TSTT = np.sum(flow * travel_time)

print("Baseline Total Travel Time:", baseline_TSTT)

# ------------------------------------------------------------
# MONTE CARLO SIMULATION
# ------------------------------------------------------------

simulation_results = []

with pd.ExcelWriter(file_name, mode='w', engine='openpyxl') as writer:

    for i in range(n_simulations):

        # Reset all capacities before every simulation
        h.SetMulti( Visum.Net.Links, "CAPPRT",original_capacity)

        failure_probability = []
        failed_links = []

        # ----------------------------------------------------
        # Assign probability of failure to each link
        # ----------------------------------------------------

        for l in range(len(from_node)):

            if link_type[l] == 91:
                # High-risk link
                p = np.random.beta(high_alpha, high_beta)

            elif link_type[l] == 92:
                # Low-risk link
                p = np.random.beta(low_alpha, low_beta)

            else:
                p = 0.0

            failure_probability.append(p)

            # Bernoulli failure event
            if np.random.rand() < p:

                disrupted_link = Visum.Net.Links.ItemByKey(
                    from_node[l][1],
                    to_node[l][1]
                )

                disrupted_link.SetAttValue(
                    "CAPPRT",
                    original_capacity[l] * failed_capacity
                )

                failed_links.append(l)

        # ----------------------------------------------------
        # RUN VISUM AFTER FAILURES
        # ----------------------------------------------------

        Proc.Execute()

        # ----------------------------------------------------
        # READ NETWORK PERFORMANCE
        # ----------------------------------------------------

        link_id = np.array(h.GetMulti(Visum.Net.Links, "No"))

        length = np.array(h.GetMulti(Visum.Net.Links, "LENGTH"))

        capacity = np.array(h.GetMulti(Visum.Net.Links, "CAPPRT"))

        v_cur = np.array(h.GetMulti(Visum.Net.Links, "VCur_PrTSys(C)"))

        flow = np.array(h.GetMulti(Visum.Net.Links, "VolVehPrT(AP)"))

        travel_time = np.array(h.GetMulti(Visum.Net.Links, "TCur_PrTSys(C)"))

        free_flow_speed = np.array(h.GetMulti(Visum.Net.Links, "V0_PrTSys(C)"))

        
        # Flow × Travel Time
        FaCa = flow * travel_time

        # Total system travel time
        TSTT = np.sum(FaCa)

        # ----------------------------------------------------
        # SERVICEABILITY
        # ----------------------------------------------------

        serviceability = baseline_TSTT / TSTT

        # ----------------------------------------------------
        # VULNERABILITY
        # ----------------------------------------------------

        vulnerability = 1 - serviceability

        # ----------------------------------------------------
        # STORE SIMULATION RESULT
        # ----------------------------------------------------

        simulation_results.append({
            "Simulation": i + 1,
            "Baseline TSTT": baseline_TSTT,
            "Disrupted TSTT": TSTT,
            "Serviceability": serviceability,
            "Vulnerability": vulnerability,
            "Failed Links": len(failed_links)
        })

        # ----------------------------------------------------
        # LINK-LEVEL DATA
        # ----------------------------------------------------

        data = {
            "Simulation": [i + 1] * len(link_id),
            "Link ID": link_id,
            "From Node": np.array(from_node)[:, 1],
            "To Node": np.array(to_node)[:, 1],
            "Link Type": link_type,
            "Length [km]": length,
            "Capacity": capacity,
            "Failure Probability": failure_probability,
            "Free Flow Speed [km/h]": free_flow_speed,
            "Flow [veh]": flow,
            "Speed [km/h]": v_cur,
            "Travel Time [s]": travel_time,
            "Time on network": FaCa
        }

        df = pd.DataFrame(data)

        # One sheet containing the link results
        # for this simulation
        df.to_excel(writer,sheet_name=f"sim_{i+1}",index=False)

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    results_df = pd.DataFrame(simulation_results)

    results_df.to_excel(writer,sheet_name="MonteCarlo_Summary",index=False)

# ------------------------------------------------------------
# RESET NETWORK
# ------------------------------------------------------------

h.SetMulti(
    Visum.Net.Links,
    "CAPPRT",
    original_capacity
)

print("\nMonte Carlo simulation completed.")
print("Number of simulations:", n_simulations)
print("Baseline TSTT:", baseline_TSTT)
print(
    "Mean Serviceability:",
    results_df["Serviceability"].mean()
)
print(
    "Mean Vulnerability:",
    results_df["Vulnerability"].mean()
)

Proc = None
Visum = None