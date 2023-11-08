import numpy as np
import pandas as pd
import openai
from pptx import Presentation
from docx import Document

# Replace 'your_api_key_here' with your actual OpenAI API key
openai.api_key = ''

from openai import OpenAI
client = OpenAI(api_key = '')
# Function to reword text using GPT-4
#def reword_text_with_gpt4(text):
#    try:
#        # Use the OpenAI API to get a response
#        response = openai.chat.completions.create(
#            model="gpt-3.5-turbo-1106",  # Replace with the appropriate GPT-4 model when available
#            messages=[
#          {"role": "system", "content": "Reword the following text to be clear and concise:\n\n" + text}
#                ]
#            #prompt="Reword the following text to be clear and concise:\n\n" + text,
#            #max_tokens=60  # Adjust max tokens as needed
#        )
#        return response.choices[0].text.strip()
#    except Exception as e:  # Catch a general exception
#        print(f"An error occurred: {e}")
#        return text  # Return the original text if an error occurs

def reword_text_with_gpt4(text, audience_type, shape_type):
    if len(text) == 0:
        return text
    try:
        if shape_type =="Title 1":
            title_prompt = " This is a title of a slide so keep it to less than 8 words."
        else:
            title_prompt =""
        # Use the OpenAI API to get a response
        response = openai.chat.completions.create(
            model="gpt-4-0613",  # Replace with the appropriate model
            messages=[
                {
                    "role": "system", 
                    "content": (
                        "You are an assistant that rewords sentences to be clear and concise. Your output will be no longer than the input in length." +
                        " I am presenting to " + audience_type + ", so make it suitable for this audience." + title_prompt

                                )
                    },
                {"role": "user", "content": text}
            ]
        )
        # Assuming the last message in the list will be the assistant's response
        return str(response.choices[0].message.content)#.text.strip()
    except Exception as e:  # Catch a general exception
        print(f"An error occurred: {e}")
        return text  # Return the original text if an error occurs

def create_exec_summary(text, audience_type):
    try:
        # Use the OpenAI API to get a response
        response = openai.chat.completions.create(
            model="gpt-4-0613",  # Replace with the appropriate model
            messages=[
                {
                    "role": "system", 
                    "content": (
                        "You are an assistant that creates executive summaries." +
                        " I am presenting to " + audience_type + ", so make it suitable for this audience. Produce a simple five bullet point summary"

                                )
                    },
                {"role": "user", "content": text}
            ]
        )
        # Assuming the last message in the list will be the assistant's response
        return str(response.choices[0].message.content)#.text.strip()
    except Exception as e:  # Catch a general exception
        print(f"An error occurred: {e}")
        return text  # Return the original text if an error occurs
def create_ppt_feedback(text, audience_type):
    try:
        # Use the OpenAI API to get a response
        response = openai.chat.completions.create(
            model="gpt-4-0613",  # Replace with the appropriate model
            messages=[
                {
                    "role": "system", 
                    "content": (
                        "You are an assistant that provides feedback for powerpoint presentations." +
                        " I am presenting to " + audience_type + ", so make it suitable for this audience." +
                       " Produce feedback on the following powerpoint contents and tell me how to improve it."

                                )
                    },
                {"role": "user", "content": text}
            ]
        )
        # Assuming the last message in the list will be the assistant's response
        return str(response.choices[0].message.content)#.text.strip()
    except Exception as e:  # Catch a general exception
        print(f"An error occurred: {e}")
        return text  # Return the original text if an error occurs

# Function to process the PowerPoint file
def process_presentation(
    input_file_path, 
    output_file_path, 
    audience_type,
    executive_summary_slide = False):
    # Load the presentation
    prs = Presentation(input_file_path)
    
    # Iterate through each slide and each text box
    all_text = ""
    for slide_number, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                original_text = shape.text
                # Reword the text using GPT-4
                reworded_text = reword_text_with_gpt4(original_text, audience_type, shape.name)

                # Replace the original text with the reworded text
                shape.text = reworded_text
                all_text = all_text + ". " + original_text
        print(f"Processed slide {slide_number + 1}")
    
    if executive_summary_slide == True:
        print("Creating an executive summary slide")
        slide_layout = prs.slide_layouts[1]
        slide_exec = prs.slides.add_slide(slide_layout)
        slide_exec.placeholders[0].text = "Executive Summary"
        slide_exec.placeholders[1].text = create_exec_summary(all_text, audience_type)

        feedback = create_ppt_feedback(all_text, audience_type)
        f = open("feedback.txt", "a")
        f.write(feedback)
        f.close()
    # Save the presentation
    prs.save(output_file_path)
    print(f"Presentation saved to {output_file_path}")



def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

def document_to_ppt(doc,ppt, audience_type, slide_number):

    
    soup = bs(open(doc).read())
    [s.extract() for s in soup(['style', 'script'])]
    tmpText = soup.get_text()
    text = "".join("".join(tmpText.split('\t')).split('\n')).strip()
    print(text)

    try:
        # Use the OpenAI API to get a response
        response = openai.chat.completions.create(
            model="gpt-4-0613",  # Replace with the appropriate model
            messages=[
                {
                    "role": "system", 
                    "content": (
                        "You are an assistant that provides powerpoints form word document inputs." +
                        " I am presenting to " + audience_type + ", so convert this document into a powerpoint suitable for this audience." +
                       "Ensure it has " + str(slide_number + 1) + " slides. The first slide is a title slide and the other slides are in the form of title and contents." +
                       "Provide  the answer as a dictionary for python of the form {1:{'title':'title','subtitle: 'appropriate subtitle'},2:{'title':'title','contents: 'bullet pointed contents'},3:{'title':'title','contents: 'bullet pointed contents'}... etc}"

                                )
                    },
                {"role": "user", "content": text}
            ]
        )
        # Assuming the last message in the list will be the assistant's response
        ppt_dict =  str(response.choices[0].message.content)#.text.strip()
        print(ppt_dict)
        prs = Presentation()

        title_layout = prs.slide_layouts[0]
        normal_layout = prs.slide_layouts[1]
        title_slide = prs.slides.add_slide(title_layout)
        title_slide.placeholders[0].text = ppt_dict[1]["title"]
        title_slide.placeholders[1].text = ppt_dict[1]["subtitle"]

        for i in range(2,slide_numer +1):
            slide = prs.slides.add_slide(title_layout)
            slide.placeholders[0].text = ppt_dict[i]["title"]
            slide.placeholders[1].text = ppt_dict[i]["contents"]

        # Save the presentation
        prs.save(ppt)
        print(f"Presentation saved to {ppt}")

    except Exception as e:  # Catch a general exception
        print(f"An error occurred: {e}")

    
# Example usage
wd = "C://Users//Administrator//Documents//GitHub//genai_group1"
import os
os.chdir(wd)
input_file_path =  "GPTB4.pptx"
output_file_path = "GPTA4"
output_file_path_technical = "GPTA4_tech.pptx"
output_file_path_baby = "GPTA4_babies.pptx"
audience_type1 = "a technical audience"
audience_type2 = "a bunch of five year olds who like thomas the tank engine"
document_to_ppt("Singapore Day 1.docx","singapore.pptx",audience_type2, 3)
#process_presentation(input_file_path, output_file_path_technical, audience_type1,executive_summary_slide = True)
#process_presentation(input_file_path, output_file_path_baby, audience_type2)
