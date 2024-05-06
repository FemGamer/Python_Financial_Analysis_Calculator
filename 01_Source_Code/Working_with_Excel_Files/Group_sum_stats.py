##file for doing descriptive statistcs together for combined ratios dataset
import pandas as pd
import matplotlib.pyplot as plt

Ratios = "D:/PY_FemGamer/Thesis/04_Results/Ratios_Dataset.xlsx" ##this file was manully created in excel, by coming all ratios in 1 worksheet
sheet_activate= "Dataset"


df= pd.read_excel(Ratios, sheet_name=sheet_activate, index_col=0)

summary_stats= df.describe() ##function for performing descriptive statistics
graph_df= summary_stats
graph = graph_df.plot.bar()

save_graph= graph.get_figure()
output_file= "D:/PY_FemGamer/Thesis/04_Results/Group_Stats.xlsx" ##output saved in the excel file

with pd.ExcelWriter(output_file) as writer:
    summary_stats.to_excel(writer, sheet_name="Summary_Statistics", index=['count', 'mean', 'std', 'min', '25%', '50%', '75%', 'max'])

    ###visualising the statistical outcome va graphs
    plt.figure(figsize=(10, 6))
    plot_df= summary_stats.plot.bar(ax=plt.gca())
    plt.title('Summary Statistics')
    plt.xlabel('Metrics')
    plt.ylabel('Values')
    plt.xticks(rotation=45)
    plt.tight_layout()
    save_graph = plt.gcf()
    save_graph.savefig("D:/PY_FemGamer/Thesis/04_Results/graph_1_2.png")

plt.show()
print("Results run successful!")