import requests
from bs4 import BeautifulSoup
import openpyxl
from unidecode import unidecode
import openai
import json
import re
from wordpress_xmlrpc import Client, WordPressPost
from wordpress_xmlrpc.methods.posts import NewPost

# setting variables
workbook = openpyxl.load_workbook('C:\\Users\\isamc\\OneDrive\\Documents\\w\\Bike Subjects for Blog.xlsx')
sheet = workbook['Bikes']
row = 2

# The loop
while row < 30:
    wbproduct_name = "A" + str(row)
    wbproduct_category = "B" + str(row)
    wbproduct_asin = "C" + str(row)
    wbst1 = "D" + str(row)
    wbst2 = "E" + str(row)
    wbst3 = "F" + str(row)
    wbproduct_mkey = "G" + str(row)
    wbproduct_skey = "H" + str(row)
    wbproduct_prompt = "I" + str(row)
    wbbullet_prompt = "K" + str(row)
    wbmeta = "L" + str(row)
    wbintro = "M" + str(row)
    wbassembly = "N" + str(row)
    wbdesign = "O" + str(row)
    wbcomfort = "P" + str(row)
    wbperfor = "Q" + str(row)
    wbverdict = "R" + str(row)
    wbalt = "S" + str(row)
    product_name = sheet[wbproduct_name].value
    product_category = sheet[wbproduct_category].value
    product_asin = sheet[wbproduct_asin].value
    url1 = sheet[wbst1].value
    url2 = sheet[wbst2].value
    url3 = sheet[wbst3].value
    product_mkey = sheet[wbproduct_mkey].value
    product_skey = sheet[wbproduct_skey].value
    product_prompt = sheet[wbproduct_prompt].value
    bullet_prompt = sheet[wbbullet_prompt].value
    meta_prompt = sheet[wbmeta].value
    intro_prompt = sheet[wbintro].value
    assembly_prompt = sheet[wbassembly].value
    design_prompt = sheet[wbdesign].value
    comfort_prompt = sheet[wbcomfort].value
    perfo_prompt = sheet[wbperfor].value
    verdict_prompt = sheet[wbverdict].value
    alt_prompt = sheet[wbalt].value

    # article 1
    if url1 is not None:    
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}
            response = requests.get(url1, headers=headers)
            soup = BeautifulSoup(response.text, 'html.parser')
            artc1_lst = []
            for paragraph in soup.find_all('p'):
                article = unidecode(paragraph.text)
                artc1_lst.append(article)
        except:
            print("Failed to retrieve the webpage 1")
        article1 = '\n'.join(artc1_lst)
    elif url1 is None:
        article1 = ""
    # article 2
    if url2 is not None:
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}
            response = requests.get(url2, headers=headers)
            soup = BeautifulSoup(response.text, 'html.parser')
            artc2_lst = []
            for paragraph in soup.find_all('p'):
                article = unidecode(paragraph.text)
                artc2_lst.append(article)
        except:
            print("Failed to retrieve the webpage 2")
        article2 = '\n'.join(artc2_lst)
    elif url2 is None:
        article2 = ""
    # article 3
    if url3 is not None:
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}
            response = requests.get(url3, headers=headers)
            soup = BeautifulSoup(response.text, 'html.parser')
            artc3_lst = []
            for paragraph in soup.find_all('p'):
                article = unidecode(paragraph.text)
                artc3_lst.append(article)
        except:
            print("Failed to retrieve the webpage 3")
        article3 = '\n'.join(artc3_lst)
    elif url3 is None:
        article3 = ""

    # chat_gpt
    api_key = "***********"
    openai.api_key = api_key

    # prompt url1
    f_prompt1 = product_prompt + "\n" + "--------------------" + "\n" + article1
    response1 = openai.ChatCompletion.create(
        temperature = 0.5,
        model="gpt-3.5-turbo-16k",
        messages=[{"role": "system", "content": "You are a helpful assistant that summarizes articles."}, {"role": "user", "content": f_prompt1}],
        timeout = 20000
    )
    try:
        rslt1 = response1.choices[0].message.content

    except:
        print("Error: Failed to generate a summary for the prompt 1. Check your API key and request.")

    # prompt url2
    f_prompt2 = product_prompt + "\n" + "--------------------" + "\n" + article2
    response2 = openai.ChatCompletion.create(
        temperature = 0.5,
        model="gpt-3.5-turbo-16k",
        messages=[{"role": "system", "content": "You are a helpful assistant that summarizes articles."}, {"role": "user", "content": f_prompt2}],
        timeout = 20000
    )
    try:
        rslt2 = response2.choices[0].message.content

    except:
        print("Error: Failed to generate a summary for the prompt 2. Check your API key and request.")

    # prompt url3
    f_prompt3 = product_prompt + "\n" + "--------------------" + "\n" + article3
    response3 = openai.ChatCompletion.create(
        temperature = 0.5,
        model="gpt-3.5-turbo-16k",
        messages=[{"role": "system", "content": "You are a helpful assistant that summarizes articles."}, {"role": "user", "content": f_prompt3}],
        timeout = 20000
    )
    try:
        rslt3 = response3.choices[0].message.content

    except:
        print("Error: Failed to generate a summary for the prompt 3. Check your API key and request.")

    # bullet points
    bullet = bullet_prompt + "\n" + rslt1 + "\n" + rslt2 + "\n" + rslt3

    bullet_response = openai.ChatCompletion.create(
        temperature = 0.5,
        model="gpt-3.5-turbo-16k",
        messages=[{"role": "system", "content": "You are a helpful assistant that summarizes articles."}, {"role": "user", "content": bullet}],
        timeout = 20000
    )
    try:
        bullet_points = bullet_response.choices[0].message.content

    except:
        print("Error: Failed to generate a summary for the prompt 3. Check your API key and request.")

    # making a dictionary
    bike_dict = json.loads(bullet_points)

    # Intro section
    intro = []
    try:
        for header, value in bike_dict.items():
            if re.match(re.compile(r'^intro', re.IGNORECASE), header):
                intro.append(value)
    except:
        print("Intro section not found")

    intro_text = ""
    for item in intro[0]:
        intro_text += f"{item}\n"

    # Assembly section
    assemb = []
    try:
        for header, value in bike_dict.items():
            if re.match(re.compile(r'^assemb', re.IGNORECASE), header):
                assemb.append(value)
    except:
        print("Assembly section not found")

    assem_text = ""
    for item in assemb[0]:
        assem_text += f"{item}\n"

    # Design section
    design = []
    try:
        for header, value in bike_dict.items():
            if re.match(re.compile(r'^desig', re.IGNORECASE), header):
                design.append(value)
    except:
        print("Design section not found")

    design_text = ""
    for item in design[0]:
        design_text += f"{item}\n"

    # Comfort section
    comf = []
    try:
        for header, value in bike_dict.items():
            if re.match(re.compile(r'^comfo', re.IGNORECASE), header):
                comf.append(value)
    except:
        print("Comfort section not found")

    comf_text = ""
    for item in comf[0]:
        comf_text += f"{item}\n"
    
    # Performance section
    perf = []
    try:
        for header, value in bike_dict.items():
            if re.match(re.compile(r'^perfor', re.IGNORECASE), header):
                perf.append(value)
    except:
        print("Performance section not found")

    perfo_text = ""
    for item in perf[0]:
        perfo_text += f"{item}\n"

    # Verdict section
    verdict = []
    try:
        for header, value in bike_dict.items():
            if re.match(re.compile(r'^verdict', re.IGNORECASE), header):
                verdict.append(value)
    except:
        print("Verdict section not found")

    verdict_text = ""
    for item in verdict[0]:
        verdict_text += f"{item}\n"

    # Alternatives section
    alter = []
    try:
        for header, value in bike_dict.items():
            if re.match(re.compile(r'^alternati', re.IGNORECASE), header):
                alter.append(value)
    except:
        print("Alternatives section not found")

    alter_text = ""
    for item in alter[0]:
        alter_text += f"{item}\n"

    # meta_prompt
    p_meta = meta_prompt
    response4 = openai.ChatCompletion.create(
        temperature = 0.7,
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant that writes articles."}, {"role": "user", "content": p_meta}],
        timeout = 20000
    )
    try:
        meta_section = response4.choices[0].message.content

    except:
        print("Error: Failed to generate the meta section. Check your API key and request.")

    # intro_prompt
    p_intro = intro_prompt + "\n" + intro_text
    response5 = openai.ChatCompletion.create(
        temperature = 0.7,
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant that writes articles."}, {"role": "user", "content": p_intro}],
        timeout = 20000
    )
    try:
        intro_section = response5.choices[0].message.content

    except:
        print("Error: Failed to generate the intro section. Check your API key and request.")

    # assembly_prompt
    p_assem = assembly_prompt + "\n" + assem_text
    response6 = openai.ChatCompletion.create(
        temperature = 0.7,
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant that writes articles."}, {"role": "user", "content": p_assem}],
        timeout = 20000
    )
    try:
        assembly_section = response6.choices[0].message.content

    except:
        print("Error: Failed to generate the assembly section. Check your API key and request.")

    # design_prompt
    p_design = design_prompt + "\n" + design_text
    response7 = openai.ChatCompletion.create(
        temperature = 0.7,
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant that writes articles."}, {"role": "user", "content": p_design}],
        timeout = 20000
    )
    try:
        design_section = response7.choices[0].message.content

    except:
        print("Error: Failed to generate the design section. Check your API key and request.")

    # comfort_prompt
    p_comf = comfort_prompt + "\n" + comf_text
    response8 = openai.ChatCompletion.create(
        temperature = 0.7,
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant that writes articles."}, {"role": "user", "content": p_comf}],
        timeout = 20000
    )
    try:
        comfort_section = response8.choices[0].message.content

    except:
        print("Error: Failed to generate the comfort section. Check your API key and request.")

    # performance_prompt
    p_perfo = perfo_prompt + "\n" + perfo_text
    response9 = openai.ChatCompletion.create(
        temperature = 0.7,
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant that writes articles."}, {"role": "user", "content": p_perfo}],
        timeout = 20000
    )
    try:
        performance_section = response9.choices[0].message.content

    except:
        print("Error: Failed to generate the performance section. Check your API key and request.")   
    
    # verdict_prompt
    p_verd = verdict_prompt + "\n" + verdict_text
    response10 = openai.ChatCompletion.create(
        temperature = 0.7,
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant that writes articles."}, {"role": "user", "content": p_verd}],
        timeout = 20000
    )
    try:
        verdict_section = response10.choices[0].message.content

    except:
        print("Error: Failed to generate the verdict section. Check your API key and request.")
    
    # alternatives_prompt
    p_alt = alt_prompt + "\n" + alter_text
    response11 = openai.ChatCompletion.create(
        temperature = 0.7,
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant that writes articles."}, {"role": "user", "content": p_alt}],
        timeout = 20000
    )
    try:
        alternatives_section = response11.choices[0].message.content

    except:
        print("Error: Failed to generate the alternatives section. Check your API key and request.")

    # Wordpress
    wordpress_url = '********'
    username = '********'
    password = '********'

    client = Client(wordpress_url, username, password)

    post = WordPressPost()
    post.content = meta_section + product_asin + "\n" + intro_section + assembly_section + design_section + comfort_section + performance_section + verdict_section + alternatives_section
    post.post_status = 'draft'

    # Publish the post
    client.call(NewPost(post))
    print("article " + str(row) + " published.")
    row = row + 1
print("Done.")