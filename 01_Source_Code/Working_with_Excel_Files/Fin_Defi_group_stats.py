import pandas as pd
import matplotlib.pyplot as plt

Ratios = "D:/PY_FemGamer/Thesis/04_Results/Ratios_Dataset_2.xlsx"
sheet_activate= "Total"


df= pd.read_excel(Ratios, sheet_name=sheet_activate, index_col=0)

summary_stats= df.describe()
graph_df= summary_stats
graph = graph_df.plot.bar()

save_graph= graph.get_figure()
output_file= "D:/PY_FemGamer/Thesis/04_Results/Ratios_Stats_2.xlsx"

with pd.ExcelWriter(output_file) as writer:
    summary_stats.to_excel(writer, sheet_name="Summary_Statistics", index=['count', 'mean', 'std', 'min', '25%', '50%', '75%', 'max'])
    #save_graph.savefig("D:/PY_FemGamer/Thesis/04_Results/graph_02.png")
    
    plt.figure(figsize=(10, 6))
    plot_df= summary_stats.plot.bar(ax=plt.gca())
    plt.title('Summary Statistics')
    plt.xlabel('Metrics')
    plt.ylabel('Values')
    plt.xticks(rotation=45)
    plt.tight_layout()
    save_graph = plt.gcf()
    save_graph.savefig("D:/PY_FemGamer/Thesis/04_Results/graph_03.png")

plt.show()
print("Results run successful!")