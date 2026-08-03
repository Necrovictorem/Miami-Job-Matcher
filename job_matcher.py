import os #lets Python read environment variables (this is how the script will read your Adzuna App ID/Key without them being written in the code itself).
import json #reads/writes JSON data; we use it to load skills.json.
import requests #the standard Python library for making HTTP calls. This is what actually talks to the Adzuna API. It's not built into Python by default, so it needs to be installed (that's what requirements.txt)
from datetime import date
import openpyxl #writes real .xlsx Excel files. Also not built-in, also goes in requirements.txt.

def fetch_jobs(app_id, app_key, what="engineer", where="Miami", pages=3): #Function that will be called multiple times to search for jobs
    all_jobs = []
    for page in range(1, pages + 1):
        url = f"https://api.adzuna.com/v1/api/jobs/us/search/{page}"
        params = {
            "app_id": app_id,
            "app_key": app_key,
            "what": what,
            "where": where,
            "results_per_page": 50,
            "content-type": "application/json"
        }
        response = requests.get(url, params=params)
        response.raise_for_status() #checks whether the API returned an error (like a 401 for bad credentials or 429 for rate-limited) and throws a Python exception if so, so the script fails loudly instead of silently returning garbage.
        data = response.json() #parses the response body from raw text into a Python dictionary.
        all_jobs.extend(data.get("results", [])) #adds each item in a list individually to all_jobs; append would nest the whole page's list as one single item. We want a flat list of individual job postings, so extend is correct here.
    return all_jobs

def score_job(job, skills_data):
    text = (job.get("title", "") + " " + job.get("description", "")).lower()
    score = 0
    matched = []

    all_terms = skills_data["core_skills"] + skills_data["tools"] + skills_data["domains"]
    weights = skills_data.get("weight_overrides", {})

    for term in all_terms:
        if term.lower() in text:
            score += weights.get(term, 1)
            matched.append(term)

    return score, matched

def build_spreadsheet(jobs, skills_data, min_score=3, filename=None):
    if filename is None:
        filename = f"output/matches_{date.today()}.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matches"
    ws.append(["Score", "Title", "Company", "Location", "Matched Skills", "URL"])

    for job in jobs:
        score, matched = score_job(job, skills_data)
        if score >= min_score:
            ws.append([
                score,
                job.get("title", ""),
                job.get("company", {}).get("display_name", ""),
                job.get("location", {}).get("display_name", ""),
                ", ".join(matched),
                job.get("redirect_url", "")
            ])

    wb.save(filename)
    return filename


def main():
    app_id = os.environ["ADZUNA_APP_ID"]
    app_key = os.environ["ADZUNA_APP_KEY"]

    with open("data/skills.json") as f:
        skills_data = json.load(f)

    search_terms = ["engineer", "technical consultant", "senior technical consultant"]
    jobs = []
    for term in search_terms:
        jobs.extend(fetch_jobs(app_id, app_key, what=term))

    filename = build_spreadsheet(jobs, skills_data)
    print(f"Saved {len(jobs)} jobs checked, results in {filename}")

if __name__ == "__main__":
    main()