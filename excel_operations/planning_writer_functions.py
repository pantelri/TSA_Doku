import matplotlib.pyplot as plt
import seaborn as sns
import os
from matplotlib.ticker import FuncFormatter

def format_y_axis(x, p):
    if x >= 1e6:
        return f'{x/1e6:.0f}M'
    elif x >= 1e3:
        return f'{x/1e3:.0f}K'
    else:
        return f'{x:.0f}'

def create_parameter_plot(data, param, output_dir):
    plt.figure(figsize=(12, 6))
    
    if param.startswith("index_"):
        # Connected Scatter Plot für Index-Parameter
        sns.scatterplot(data=data, x="Date", y=param, hue="Period", palette={"Audit Period": "orange", "Prior Period": "gray"})
        plt.plot(data["Date"], data[param], color="blue", alpha=0.5)
    else:
        # Barplot für andere Parameter
        sns.barplot(data=data, x="Date", y=param, hue="Period", palette={"Audit Period": "orange", "Prior Period": "gray"})
    
    plt.title(f"Development of {param}")
    plt.xticks(rotation=45)
    
    # Formatierung der y-Achse
    plt.gca().yaxis.set_major_formatter(FuncFormatter(format_y_axis))
    
    plt.tight_layout()
    
    # Speichern des Plots
    filename = os.path.join(output_dir, f"{param}_plot.png")
    plt.savefig(filename)
    plt.close()
