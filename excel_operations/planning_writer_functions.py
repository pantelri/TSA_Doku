import matplotlib.pyplot as plt
import seaborn as sns
import os

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
    plt.tight_layout()
    
    # Speichern des Plots
    filename = os.path.join(output_dir, f"{param}_plot.png")
    plt.savefig(filename)
    plt.close()
