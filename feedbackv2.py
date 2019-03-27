# -*- coding: utf-8 -*-
"""
Created on Mon Mar 18 20:52:54 2019

@author: Jan-Gunther Gosselke
"""
import os
import pandas as pd
import numpy as np

path = "C:/Users/Jan-Gunther Gosselke/Downloads/OPD_Feedback"
os.chdir(path)



#-------------------------Preparation
judges = pd.read_excel(io = "Input/Judges.xlsx", sheet = "Tabelle1")
chairs = pd.read_excel(io = "Input/Hauptjuroren.xlsx", sheet = "Tabelle1")
wings = pd.read_excel(io = "Input/Nebenjuroren.xlsx", sheet = "Tabelle1")

feedback_by_chair = pd.read_excel(io = "Input/Hauptfeedback.xlsx", index_col = [2,3]).drop(["Zeitstempel"], axis = 1)
feedback_by_chair = feedback_by_chair.rename(index=str, columns={"Feedbackender Hauptjuror" : "by_haupt",
                                                                 "Gefeedbackter Nebenjuror" : "judge",
                                                                 "H1: Die Punktedifferenzen zwischen den einzelnen Juroren waren?" : "ergebnisfindung_H",
                                                                 "H2: Die Kompetenz des Jurors war aus meiner Sicht?" : "h2",
                                                                 "H3: Die vom Juror vergebenen Punktzahlen und Begründungen sind stimmig?" : "h3",
                                                                 "H4: Trägt der Juror konstruktiv zur Ergebnisfindung bei?" : "h4",
                                                                 "H5: Würdest du dem Juror zutrauen, in der nächsten Runde selbst das Feedback zu geben?" : "h5",
                                                                 "H6: Würdest du dem Juror generell zutrauen, eine Breakrunde zu jurieren?" : "h6"})
feedback_by_wing = pd.read_excel(io = "Input/Nebenfeedback.xlsx", index_col = [2,3]).drop(["Zeitstempel"], axis = 1)
feedback_by_wing = feedback_by_wing.rename(index=str, columns={"Feedbackender Nebenjuror" : "by_neben",
                                                               "Gefeedbackter Hauptjuror" : "judge",
                                                               "N1: Die Punktedifferenzen zwischen den einzelnen Juroren waren" : "ergebnisfindung_N",
                                                               "N2: Allgemeine Kompetenz des Jurors" : "n2",
                                                               "N3: Sind Punktzahlen und Begründungen stimmig?" : "n3",
                                                               "N4: Wie konstruktiv hat der Juror die Diskussion moderiert?" : "n4",
                                                               "N5: Wie hoch war die Qualität des Feedback?" : "n5"})

feedback_by_team = pd.read_excel(io = "Input/Teamfeedback.xlsx", index_col = [2,3]).drop(["Zeitstempel"], axis = 1)
feedback_by_team = feedback_by_team.rename(index=str, columns={"Feedbackendes Team" : "by_team",
                                                               "Gefeedbackter Hauptjuror" : "judge",
                                                               "T1: Wie hoch war die Kompetenz des Jurors?" : "t1",
                                                               "T2: Gemessen an der gegnerischen Punktzahl war unsere eigene Punktzahl?" : "t2",
                                                               "T3: Punkteniveau der Debatte nachvollziehbar" : "t3",
                                                               "T4: Plausible Erklärung der Teamkategorien" : "t4",
                                                               "T5: Allgemeine Feedbackqualität" : "t5"})

feedback_by_freespeaker = pd.read_excel(io = "Input/Freifeedback.xlsx", index_col = [2,3]).drop(["Zeitstempel"], axis = 1)
feedback_by_freespeaker = feedback_by_freespeaker.rename(index=str, columns={"Feedbackender Redner" : "by_frei",
                                                               "Gefeedbackter Hauptjuror" : "judge",
                                                               "F1: Wie hoch war die Kompetenz des Jurors?" : "f1",
                                                               "F2: Nach Hören des Feedbacks war meine eigene Punktzahl?" : "f2",
                                                               "F3: Punkteniveau der Debatte nachvollziehbar" : "f3",
                                                               "F4: Allgemeine Feedbackqualität" : "f4"})

total_df = pd.concat([feedback_by_chair, feedback_by_wing, feedback_by_team, feedback_by_freespeaker], axis = 0)
total_df = total_df.merge(judges, how = "left", left_on = "judge", right_on = "judge")
total_df = total_df.merge(chairs, how = "left", left_on = "by_haupt", right_on = "by_haupt")
total_df = total_df.merge(wings, how = "left", left_on = "by_neben", right_on = "by_neben")
total_df["t2_centered"] = total_df["t2"] - 3 #Wert von 3 entspricht "genau richtig", kleinerer Wert entspricht "zu niedrig", höherer "zu hoch" -> zentrieren
total_df["f2_centered"] = total_df["f2"] - 3 #Wert von 3 entspricht "genau richtig", kleinerer Wert entspricht "zu niedrig", höherer "zu hoch" -> zentrieren
total_df["h5_mod"] = total_df["h5"]*1.25 #Nutzt vierer Skala auf Bogen, um Entscheidung zu erzwingen; muss daher auf 5er Skala skaliert werden
total_df["h6_mod"] = total_df["h6"]*1.25 #Nutzt vierer Skala auf Bogen, um Entscheidung zu erzwingen; muss daher auf 5er Skala skaliert werden

#----------------------------Analysis
categories = {"f1":["frei","F1_allgemeine_Kompetenz_des_HJ"],
              "f2_centered":["frei","F2_Angemessenheit_eigene_Punktzahl"],
              "f3":["frei","F3_Punkteniveau_nachvollziehbar"],
              "f4":["frei","F4_generelle_Feedbackqualitaet"],
              "ergebnisfindung_H":["haupt","H1_Punktedifferenz_Juroren"],
              "h2":["haupt","H2_allgemeine_Kompetenz_des_NJ"],
              "h3":["haupt","H3_Stimmigkeit_Punkte_Begruendungen"],
              "h4":["haupt","H4_konstruktiver_Beitrag_Ergebnisfindung"],
              "h5":["haupt","H5_Feedback_geben"],
              "h6":["haupt","H6_Breakrunde_jurieren"],
              "h5_mod":["haupt","H5_Mod_Feedback_geben"],
              "h6_mod":["haupt","H6_Mod_Breakrunde_jurieren"],
              "ergebnisfindung_N":["neben","N1_Punktedifferenz_Juroren"],
              "n2":["neben","N2_allgemeine_Kompetenz_des_HJ"],
              "n3":["neben","N3_Stimmigkeit_Punkte_Begruendungen"],
              "n4":["neben","N4_Diskussion_gut_moderiert"],
              "n5":["neben","N5_Qualitaet_des_Feedbacks"],
              "t1":["team","T1_allgemeine_Kompetenz_des_HJ"],
              "t2_centered":["team","T2_relative_Punktzahl_angemessen"],
              "t3":["team","T3_Punkteniveau_nachvollziehbar"],
              "t4":["team","T4_Teamkategorien_erklaert"],
              "t5":["team","T5_generelle_Feedbackqualitaet"]}

for category in categories:
    print(category)
    total_df[category+"_glob_m"] = total_df[category].mean()
    total_df[category+"_rank_m"] = total_df[category].groupby(total_df["judge_rank"]).transform("mean")
    total_df[category+"_"+categories[category][0]+"_m"] = total_df[category].groupby(total_df["by_"+categories[category][0]]).transform("mean")
    total_df[category+"_diff_to_all"] = total_df[category] - total_df[category+"_glob_m"]
    total_df[category+"_diff_to_rank"] = total_df[category] - total_df[category+"_rank_m"]
    total_df[category+"_diff_to_"+categories[category][0]] = total_df[category] - total_df[category+"_"+categories[category][0]+"_m"]
    category_summary = total_df.groupby("judge").agg({
            category : ["mean","count"],
            "judge_rank" : "mean",
            category+"_diff_to_all" : "mean",
            category+"_diff_to_rank" : "mean",
            category+"_diff_to_"+categories[category][0] : "mean"})
    category_summary.columns = ['{}_{}'.format(x[0], x[1]) for x in category_summary.columns]
    category_summary.sort_values(by=[category+"_mean", category+"_diff_to_rank_mean", category+"_diff_to_" + categories[category][0] + "_mean"], ascending=False)\
    .to_excel(path+"/Output/" + categories[category][0] +"/" + categories[category][1] + ".xlsx")

#-------------------------Feedback from High Quality Judges
categories_highq = {"h2":["haupt","H2_allgemeine_Kompetenz_des_NJ"],
              "h3":["haupt","H3_Stimmigkeit_Punkte_Begruendungen"],
              "h4":["haupt","H4_konstruktiver_Beitrag_Ergebnisfindung"],
              "h5":["haupt","H5_Feedback_geben"],
              "h6":["haupt","H6_Breakrunde_jurieren"],
              "n2":["neben","N2_allgemeine_Kompetenz_des_HJ"],
              "n3":["neben","N3_Stimmigkeit_Punkte_Begruendungen"],
              "n4":["neben","N4_Diskussion_gut_moderiert"],
              "n5":["neben","N5_Qualitaet_des_Feedbacks"]}

for category in categories_highq:
    cut_off = 8 if categories[category][0] == "haupt" else 5
    total_df[category+"_glob_m_hq"] = total_df[category][total_df["judge_rank"] >= cut_off].mean()
    total_df[category+"_rank_m_hq"] = total_df[category][total_df["judge_rank"] >= cut_off].groupby(total_df["judge_rank"]).transform("mean")
    total_df[category+"_"+categories[category][0]+"_m_hq"] = total_df[category][total_df["judge_rank"] >= cut_off].groupby(total_df["by_"+categories[category][0]]).transform("mean")
    total_df[category+"_diff_to_all_hq"] = total_df[category][total_df["judge_rank"] >= cut_off] - total_df[category+"_glob_m_hq"]
    total_df[category+"_diff_to_rank_hq"] = total_df[category][total_df["judge_rank"] >= cut_off] - total_df[category+"_rank_m_hq"]
    total_df[category+"_diff_to_"+categories[category][0]] = total_df[category][total_df["judge_rank"] >= cut_off] - total_df[category+"_"+categories[category][0]+"_m_hq"]
    category_summary = total_df[total_df["judge_rank"] >= cut_off].groupby("judge").agg({
            category : ["mean","count"],
            "judge_rank" : "mean",
            category+"_diff_to_all" : "mean",
            category+"_diff_to_rank" : "mean",
            category+"_diff_to_"+categories[category][0] : "mean"})
    category_summary.columns = ['{}_{}'.format(x[0], x[1]) for x in category_summary.columns]
    category_summary.sort_values(by=[category+"_mean", category+"_diff_to_rank_mean", category+"_diff_to_" + categories[category][0] + "_mean"], ascending=False)\
    .to_excel(path+"/Output/" + categories[category][0] +"/" + categories[category][1] + "_HQ.xlsx")

#--------------------------Bögen aggregiert
total_df.loc[~((total_df["f2_centered"].isna()) & (total_df["t2_centered"].isna())),"abweichung"]\
= total_df.loc[~((total_df["f2_centered"].isna()) & (total_df["t2_centered"].isna())),["f2_centered","t2_centered"]].sum(axis=1)
total_df["abweichung"] = pd.to_numeric(total_df["abweichung"])
abweichung_summary = total_df.groupby("judge").agg({
        "judge_rank" : "mean",
        "f2_centered" : "count",
        "t2_centered" : "count",
        "abweichung" : "mean"})
abweichung_summary.columns = ["Judge Rank", "Number Free Boegen", "Number Team Boegen", "Abweichung"]
abweichung_summary.sort_values(by=["Abweichung"], ascending=False).to_excel(path+"/Output/Abweichung.xlsx")

#--------------------------Nebenbögen

total_df["mn"] = total_df[["n2","n3","n4","n5"]].mean(axis = 1)
total_df["all_mn"] = total_df[["n2_diff_to_all","n3_diff_to_all","n4_diff_to_all","n5_diff_to_all"]].mean(axis = 1)
total_df["rank_mn"] = total_df[["n2_diff_to_rank","n3_diff_to_rank","n4_diff_to_rank","n5_diff_to_rank"]].mean(axis = 1)
total_df["neben_mn"] = total_df[["n2_diff_to_neben","n3_diff_to_neben","n4_diff_to_neben","n5_diff_to_neben"]].mean(axis = 1)
neben_summary = total_df.groupby("judge").agg({
        "judge_rank" : "mean",
        "mn" : "mean",
        "all_mn" : "mean",
        "rank_mn" : "mean",
        "neben_mn" : "mean",
        "n2" : "count"})
neben_summary.sort_values(by=["mn"], ascending=False).to_excel(path+"/Output/Neben.xlsx")

cut_off = 5
total_df["mn_hq"] = total_df[total_df["judge_rank"] >= cut_off][["n2","n3","n4","n5"]].mean(axis = 1)
total_df["all_mn_hq"] = total_df[total_df["judge_rank"] >= cut_off][["n2_diff_to_all","n3_diff_to_all","n4_diff_to_all","n5_diff_to_all"]].mean(axis = 1)
total_df["rank_mn_hq"] = total_df[total_df["judge_rank"] >= cut_off][["n2_diff_to_rank","n3_diff_to_rank","n4_diff_to_rank","n5_diff_to_rank"]].mean(axis = 1)
total_df["neben_mn_hq"] = total_df[total_df["judge_rank"] >= cut_off][["n2_diff_to_neben","n3_diff_to_neben","n4_diff_to_neben","n5_diff_to_neben"]].mean(axis = 1)
neben_summary_hq = total_df.groupby("judge").agg({
        "judge_rank" : "mean",
        "mn_hq" : "mean",
        "all_mn_hq" : "mean",
        "rank_mn_hq" : "mean",
        "neben_mn_hq" : "mean",
        "n2" : "count"})
neben_summary_hq.sort_values(by=["mn_hq"], ascending=False).to_excel(path+"/Output/Neben_HQ.xlsx")


#--------------------------Hauptbögen
total_df["mh"] = total_df[["h2","h3","h4","h5_mod","h5_mod"]].mean(axis = 1)
total_df["all_mh"] = total_df[["h2_diff_to_all","h3_diff_to_all","h4_diff_to_all","h5_mod_diff_to_all","h6_mod_diff_to_all"]].mean(axis = 1)
total_df["rank_mh"] = total_df[["h2_diff_to_rank","h3_diff_to_rank","h4_diff_to_rank","h5_mod_diff_to_rank","h6_mod_diff_to_rank"]].mean(axis = 1)
total_df["haupt_mh"] = total_df[["h2_diff_to_haupt","h3_diff_to_haupt","h4_diff_to_haupt","h5_mod_diff_to_haupt","h6_mod_diff_to_haupt"]].mean(axis = 1)
haupt_summary = total_df.groupby("judge").agg({
        "judge_rank" : "mean",
        "mh" : "mean",
        "all_mh" : "mean",
        "rank_mh" : "mean",
        "haupt_mh" : "mean",
        "h2" : "count"})
haupt_summary.sort_values(by=["mh"], ascending=False).to_excel(path+"/Output/Haupt.xlsx")

cut_off = 8
total_df["mh_hq"] = total_df[total_df["judge_rank"] >= cut_off][["h2","h3","h4","h5_mod","h6_mod"]].mean(axis = 1)
total_df["all_mh_hq"] = total_df[total_df["judge_rank"] >= cut_off][["h2_diff_to_all","h3_diff_to_all","h4_diff_to_all","h5_mod_diff_to_all","h6_mod_diff_to_all"]].mean(axis = 1)
total_df["rank_mh_hq"] = total_df[total_df["judge_rank"] >= cut_off][["h2_diff_to_rank","h3_diff_to_rank","h4_diff_to_rank","h5_mod_diff_to_rank","h6_mod_diff_to_rank"]].mean(axis = 1)
total_df["haupt_mh_hq"] = total_df[total_df["judge_rank"] >= cut_off][["h2_diff_to_haupt","h3_diff_to_haupt","h4_diff_to_haupt","h5_mod_diff_to_haupt","h6_mod_diff_to_haupt"]].mean(axis = 1)
haupt_summary_hq = total_df.groupby("judge").agg({
        "judge_rank" : "mean",
        "mh_hq" : "mean",
        "all_mh_hq" : "mean",
        "rank_mh_hq" : "mean",
        "haupt_mh_hq" : "mean",
        "h2" : "count"})
haupt_summary_hq.sort_values(by=["mh_hq"], ascending=False).to_excel(path+"/Output/Haupt_HQ.xlsx")


#--------------------------Freie Bögen
total_df["mf"] = total_df[["f1","f3","f4"]].mean(axis = 1)
total_df["all_mf"] = total_df[["f1_diff_to_all","f3_diff_to_all","f4_diff_to_all"]].mean(axis = 1)
total_df["rank_mf"] = total_df[["f1_diff_to_rank","f3_diff_to_rank","f4_diff_to_rank"]].mean(axis = 1)
total_df["frei_mf"] = total_df[["f1_diff_to_frei","f3_diff_to_frei","f4_diff_to_frei"]].mean(axis = 1)
frei_summary = total_df.groupby("judge").agg({
        "judge_rank" : "mean",
        "mf" : "mean",
        "all_mf" : "mean",
        "rank_mf" : "mean",
        "frei_mf" : "mean",
        "abweichung" : "mean",  #Achtung, das ist die aggregierte Abweichung von Team und freien Rednern
        "f1" : "count"}).rename(columns={'mf': 'mf_mean', 'all_mf': 'all_mf_mean', 'rank_mf': 'rank_mf_mean', 'frei_mf': 'frei_mf_mean'})
frei_summary.sort_values(by=["mf_mean"], ascending=False).to_excel(path+"/Output/Frei.xlsx")


#--------------------------Team Bögen
total_df["mt"] = total_df[["t1","t3","t4","t5"]].mean(axis = 1)
total_df["all_mt"] = total_df[["t1_diff_to_all","t3_diff_to_all","t4_diff_to_all","t5_diff_to_all"]].mean(axis = 1)
total_df["rank_mt"] = total_df[["t1_diff_to_rank","t3_diff_to_rank","t4_diff_to_rank","t5_diff_to_rank"]].mean(axis = 1)
total_df["team_mt"] = total_df[["t1_diff_to_team","t3_diff_to_team","t4_diff_to_team","t5_diff_to_team"]].mean(axis = 1)
team_summary = total_df.groupby("judge").agg({
        "judge_rank" : "mean",
        "mt" : "mean",
        "all_mt" : "mean",
        "rank_mt" : "mean",
        "team_mt" : "mean",
        "abweichung" : "mean", #Achtung, das ist die aggregierte Abweichung von Team und freien Rednern
        "t1" : "count"}).rename(columns={'mt': 'mt_mean', 'all_mt': 'all_mt_mean', 'rank_mt': 'rank_mt_mean', 'team_mt': 'team_mt_mean'})
team_summary.sort_values(by=["mt_mean"], ascending=False).to_excel(path+"/Output/Team.xlsx")

#-------------------------Redner Bögen
total_df = total_df.merge(frei_summary[["mf_mean","rank_mf_mean","frei_mf_mean"]], left_on="judge", right_index = True, how = "left")
total_df = total_df.merge(team_summary[["mt_mean","rank_mt_mean","team_mt_mean"]], left_on="judge", right_index = True, how = "left")
total_df["mredner"] = (2*total_df["mt_mean"]+total_df["mf_mean"])/3
total_df["mf_global"] = total_df["mf_mean"].mean(axis = 0, skipna = True)
total_df["mt_global"] = total_df["mt_mean"].mean(axis = 0, skipna = True)
total_df["m_global"] = (2 * total_df["mt_global"] + total_df["mf_global"])/3
total_df["total_diff_to_all"] = total_df["mredner"]-total_df["m_global"]
total_df["total_diff_to_rank"] = (2 * total_df["rank_mt_mean"] + total_df["rank_mf_mean"])/3
total_df["total_diff_to_speaker"] = (2 * total_df["team_mt_mean"] + total_df["frei_mf_mean"])/3
speaker_summary = total_df.groupby("judge").agg({
        "judge_rank" : "mean",
        "mredner" : "mean",
        "total_diff_to_all" : "mean",
        "total_diff_to_rank" : "mean",
        "t1" : "count",
        "f1" : "count"})
speaker_summary.sort_values(by=["mredner"], ascending=False).to_excel(path+"/Output/Speaker.xlsx")
