import matplotlib.pyplot as plt
from matplotlib.widgets import CheckButtons
import pandas as pd
import argparse

# Fetch the input: an excel file holding:
#       - the companies data in the sheet called "Companies Data"
#       - the weights of the metrics in the sheet called "Metrics and Weights"
parser = argparse.ArgumentParser()                                               
parser.add_argument("--file", "-f", type=str, required=True)
args = parser.parse_args()

def get_scores(df_ratios, df_weights, score_weights):

    # Initiate the DataFrame with the scores per category and a Series with the final score
    score = pd.DataFrame(columns = df_weights.index, index = df_ratios.index)
    final_score = pd.Series(index = df_ratios.index)

    # Go through every company to see their values with their corresponding weights
    for company in score.index:
        total_score = 0
        for category in score.columns:
            # Get the weighted value of the company in the category
            weighted_values = df_ratios.loc[company] * df_weights.loc[category]
            value = weighted_values.sum()
        
            # Add the value to the dataframe with scores per category
            score[category][company] = value
        
            # Add the value to the total score, with the weight of the Category
            total_score += value * score_weights[category]
     
        # Add the total score to the total_values series
        final_score[company] = total_score     
        
    return score, final_score


def create_ranking_and_rating(df_ratios,df_weights):

    # Standardize the ratios and the weights. The ratios are standardized as opposed to the highest value
    # While the weights are standardized to add up to 1
    df_ratios = df_ratios/df_ratios.max()

    sum_score = df_weights["SCORE WEIGHT"].sum()
    score_weights = df_weights["SCORE WEIGHT"]/sum_score
    df_weights = df_weights.drop(["SCORE WEIGHT"],axis=1)

    for i in df_weights.index:
        df_weights.loc[i] = df_weights.loc[i]/df_weights.loc[i].sum()


    # Now we want to get the actual scores. Multiplying the values with the weights and finding the end-scores
    scores, total_score = get_scores(df_ratios,df_weights, score_weights)

    # Create the rankings from the total_score and the scores
    final_rank = total_score.rank(ascending=False)
    ranking = scores.rank(ascending=False)

    # Insert the end score (The score created by our algorithm) and its corrosponding rank
    # Inside the table holding the valuation of the categories
    scores.insert(0, "FINAL SCORE", total_score)
    ranking.insert(0, "FINAL RANK", final_rank)

    # Lastly, sort the vendors, from best total score/rank to worst total score/rank 
    ranking = ranking.sort_values("FINAL RANK")
    scores = scores.sort_values("FINAL SCORE",ascending=False)

    return ranking, scores

class Ratios:

    def __init__(self, df_ratios, df_weights):
        # Parse the ratios and the weights. Note that there is a regular weights and a new_weights.
        # The new_weights is the one that will be adjusted in the scatter plots
        self.ratios = df_ratios
        self.weights = self.new_weights = df_weights
        
        # Create the categories from the df_weights
        self.categories = df_weights.index
        # Create the cat_act (activated categories). First activate them all
        self.cat_act = pd.Series([True]*len(self.categories), index=self.categories)

        print(self.cat_act)
        # Update categories with initialization
        self.update_categories()

    # update_categories creates the new weights based on the cat_act
    def update_categories(self, cat_act=None):

        self.new_weights = self.weights.copy()

        # Go through the cat_act items to see which categories are activated
        for index, value in self.cat_act.items():

            if value == False:
                self.new_weights = self.new_weights.drop(index)

def create_scatter(x, y):
    
    #Clear axis
    ax.cla()
    ax.scatter(y, x)

    ax.set_title('Capstone score vs Price')

    ax.set_ylabel('Price ($)')
    ax.set_xlabel('Score')
    ax.set_xlim(0,1)

    #Create labels
    n = list(y.index)      
    for i, txt in enumerate(n):
        ax.annotate(txt, (y[i],x[i]))

    # Adjust area
    plt.subplots_adjust(left=0.3, bottom=0.3, right=0.95, top=0.95)

if __name__ == "__main__":

    # Fetch the two tables and import as DataFrames
    try:
        df_ratios = pd.read_excel(args.file, sheet_name="Companies Data", header=0, index_col=0)
        df_weights = pd.read_excel(args.file, sheet_name="Metrics and Weights", header=0, index_col=0)
    except Exception as e:
        raise e

    # Split the prices from the Dataframe
    prices = df_ratios['Price']
    print(prices)
    # Remove these prices from the DataFrame as they are not being considered as a metric
    df_ratios.drop('Price', axis=1)

    # Create an instance "Ratios" (shown above), which holds the calculation of the ratios and weights of the categories that are "active"
    # It will be created with all categories on "active"
    ratio = Ratios(df_ratios,df_weights)
    
    # Create the rankings and the ratings using the ratio that was initiated 
    rankings, ratings = create_ranking_and_rating(ratio.ratios, ratio.new_weights)

    print(rankings,"\n",ratings)

    ratings["Prices"] = -prices
    # Initiate plot, using the active categories as labels
    fig, ax = plt.subplots()
    # Create the scatterplot using the original values (all categories "activated") for Price and Score
    create_scatter(ratings["Prices"],ratings['FINAL SCORE'])

    labels = list(ratio.cat_act.index)
    # Only use the first 9 letters of the labels to keep it readable inside the checkbox
    short_labels = [i[:9] for i in labels]
    # Activated categories
    activated = ratio.cat_act.values
    # Adjust height to the amount of labels
    height = 0.05 * len(labels)
    # Create the outline of the checkbox
    axCheckButton = plt.axes([0.03, 0.05, 0.15, height])
    # Create the checkbox
    chxbox = CheckButtons(axCheckButton, short_labels, activated)

    # This function is called whenever one of the checkboxes is pressed
    def adjust_scatter(short_label):
        # Get the label from the short_label
        label = labels[short_labels.index(short_label)]
        # Reverse the category activation of the adjusted label
        ratio.cat_act[label] = not ratio.cat_act[label]
        # Update the ratios and the weights in the ratios class
        ratio.update_categories()
        # Create the new rankings and ratings based on the new ratios and new weights
        rankings,ratings = create_ranking_and_rating(ratio.ratios, ratio.new_weights)
        # Print the new rankings and ratings
        print(rankings,"\n",ratings)
        # Add prices to the DataFrame to assure it is connected to the right companies
        ratings["Prices"] = -prices
        # Update the scatter plot
        create_scatter(ratings["Prices"],ratings["FINAL SCORE"])
        plt.draw()

    # Call adjust_scatter function when checkbox is pressed
    chxbox.on_clicked(adjust_scatter)

    plt.show()

