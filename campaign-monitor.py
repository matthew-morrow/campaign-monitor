from createsend import *
from decouple import config
import pandas as pd
from tqdm import tqdm

auth = {"api_key": config("api_key")}

cs = CreateSend(auth)
clients = cs.clients()

client = Client(auth, "ec87599f8ad0495cdec716c08e719c34")

page_number = 1
# Add Sent From and To dates for weekly updates to BQ
paged_campaigns = client.campaigns(page=1)
number_of_pages = paged_campaigns.NumberOfPages
output_pd = pd.DataFrame()


while page_number <= number_of_pages:
    if page_number > 1:
        paged_campaigns = client.campaigns(page=page_number)

    print(page_number)

    for cm in tqdm(paged_campaigns.Results, desc=f"Processing page: {page_number}"):
        campaign = Campaign(auth, cm.CampaignID)
        summary = campaign.summary()

        summary_dict = {
            "campaign_id": cm.CampaignID,
            "campaign_send_date": cm.SentDate,
            "campaign_name": cm.Name,
            "campaign_subject": cm.Subject,
            "campaign_url": summary.WebVersionURL,
            "campaign_recipients": summary.Recipients,
            "campaign_total_opened": summary.TotalOpened,
            "campaign_forwards": summary.Forwards,
            "campaign_likes": summary.Likes,
            "campaign_mentions": summary.Mentions,
            "campaign_clicks": summary.Clicks,
            "campaign_unsubscribes": summary.Unsubscribed,
            "campaign_bounced": summary.Bounced,
            "campaign_unique_opens": summary.UniqueOpened,
            "campaign_complaints": summary.SpamComplaints,
        }

        df_dict = pd.DataFrame([summary_dict])
        output_pd = pd.concat([output_pd, df_dict], ignore_index=True)

    page_number = page_number + 1

output_pd.to_excel("results.xlsx")
