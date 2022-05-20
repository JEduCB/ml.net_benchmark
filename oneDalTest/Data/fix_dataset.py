import pandas as pd

#df = pd.read_csv("higgs_sample.csv")
#df["target"] = df["target"].astype(int)
#df.to_csv("higgs_sample_recode.csv", float_format='{:0.5f}'.format, index=False)

test = pd.read_csv("_airline_test.csv")
test["target"] = test["target"].astype(int)
test.to_csv("airline_test.csv", float_format='{:0.5f}'.format, index=False)

train = pd.read_csv("_airline_train.csv")
train["target"] = train["target"].astype(int)
train.to_csv("airline_train.csv", float_format='{:0.5f}'.format, index=False)