import smtplib
from email.mime.multipart import MIMEMultipart   
 #email.MIMEMultipart
from email.mime.text import MIMEText   
from email.mime.base import MIMEBase	
from email import encoders	
from openpyxl import load_workbook

workbook = load_workbook('all.xlsx')   

first_sheet = workbook.get_sheet_names()[0]
worksheet = workbook.get_sheet_by_name(first_sheet) 

fromaddr = "mnews.anshulgoyal@gmail.com"
email_range = worksheet['A1601':'A1671']




for cell1 in email_range:    
	for email in cell1:
		toaddr = email.value
		
		 
		msg = MIMEMultipart()
		 
		msg['From'] = fromaddr      
		msg['To'] = toaddr			
		msg['Subject'] = " M NEWS - EXPRESS YOURSELF"   
		 
		body = "Dear Advertiser,\n\n Media Rep is working in four verticals of business : M NEWS (Online News Channel), M Consulting (Govt. Tenders), M Productions (Web Films) and M Events.\nM NEWS is online news platform and its reach is Worldwide. It covers Political, Business, Current Affairs, Entertainment,Health, Social, Cultural, Science, Technology, Brands, Environment, Animals, Youth, Fashion, Lifestyle, Children, Education and subjects Interesting for masses.We are unbiased, politically neutral as fourth pillar of democracy. \n\nWe are active on social media which has collective reach of 50 million plus people worldwide.\n\nFacebook:       https://www.facebook.com/MNEWSONLINE/\nTwitter:            https://twitter.com/MNEWS_MEDIAREP\nLinkedin:         https://www.linkedin.com/in/mnews-mediarep-5b0553160/\n                        https://www.linkedin.com/in/abhishek-verma-05a6a015/detail/recent-activity/shares/\nBloggers:         https://mnews-mediarep.blogspot.in/\nGoogle Plus:   https://plus.google.com/u/0/113674685871480581111\nInstagram :      https://www.instagram.com/mnews.mediarep/\nSnapchat :       MNEWS\nWordpress Blog:  https://wordpress.com/view/mnews893277857.wordpress.com\nTrepup Blog:        https://www.trepup.com/mnews/timeline\nPinterest :            https://in.pinterest.com/mmediarep/\nTumblr:                https://mnews-mediarep.tumblr.com/\nIssuu:                   https://issuu.com/mnews.mediarep\nDigg:                    http://digg.com/u/MNEWS\nWhatsapp:           https://chat.whatsapp.com/7o5zMlRIPBZ5yFIqzdc26i\nReddit:                 https://www.reddit.com/user/M-NEWS/\nStumbleupon:      https://www.stumbleupon.com/stumbler/M-NEWS\nAnchor:                https://anchor.fm/m-news\nWebsites :           https://www.trepup.com/mnews\n                            https://m-news.myshopify.com\n                            www.mnews.co.in    Coming Soon\n                            www.mediarep.in   Coming Soon\nYoutube :             https://www.youtube.com/channel/UCWNi_sngEays6vgGPnbE-Mg/videos\nM MAG:              Bi-Monthly English Digital News Magazine Coming Soon\nMore Platforms:  Coming Soon\n\nWe can cover your brand story, advertorial, advertisements on above media (expanding horizontally and vertically fast) and could offer attractive deals on long term associations, annual digital media plan.\n\nKindly give an estimate budget of your annual plan you wish to spend on M News Platforms so we can make a specific plan customized for your needs of press release, branding, advertising, promotion on our digital channels mentioned above.\n\nFeel  free to discuss further queries.\nFor the complete information check this pdf : https://drive.google.com/open?id=10SNIS0g3cqbtk3eiyBD3UwAyxZdlKMYu\n\nThanks and Regards,\n\nAnshul Goyal,\nManager - Digital Sales & Marketing,\nM NEWS - EXPRESS YOURSELF \n(A Media Rep Venture)\n9882925442,  8619053557\nhttps://www.trepup.com/mnews\noperationsmnews@gmail.com\n" 

		

		msg.attach(MIMEText(body, 'plain'))
		 
		#filename = "MNEWSPPT.pdf"  
		#attachment = open("MNEWSPPT.pdf", "rb") 
		 
		#part = MIMEBase('application', 'octet-stream')
		#part.set_payload((attachment).read())
		#encoders.encode_base64(part)
		#part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
		 
		#msg.attach(part)
		 
		server = smtplib.SMTP('smtp.gmail.com', 587)
		server.starttls()
		server.login(fromaddr, "swenm@123")     
		text = msg.as_string()
		server.sendmail(fromaddr, toaddr, text)
server.quit()     
