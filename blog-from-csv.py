#make sure your CSV file has "Title" at the top

import subprocess
import importlib
import sys

def install_package(package_name):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

REQUIRED_PACKAGES = ["openai", "tqdm", "colorama", "tkinter", "python-docx","python-decouple"]

for package in REQUIRED_PACKAGES:
    try:
        importlib.import_module(package)
    except ImportError:
        install_package(package)

import openai
import os
import csv
from docx import Document
from tqdm import tqdm
from colorama import init, Fore, Style
import tkinter as tk
from tkinter import filedialog
from decouple import config

# Initialize colorama
init()

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

clear_screen()



openai.api_key = config('OPENAI_API_KEY')


def select_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title="Select CSV file", filetypes=[("CSV files", "*.csv")])
    return file_path

def load_data_from_csv(file_path):
    with open(file_path, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        return list(reader)

def generate_blog_content(title: str):
    messages = []
    messages.append({"role": "user", "content": f"""Create a blog post about '{title}'. Everything after this sentence is the information you need to know about Sofrid va Homepage/Introduction Section:

#1 Rated in 2023
Sofrid Vacuum Pro

Powerful Suction: Cleans hair, food residue, and small debris.
Versatile Cleaning: Suitable for both dry and wet cleaning, with various attachments.
Cordless and Lightweight: Easy to maneuver and access hard-to-reach areas.
Fast Charging: Quick charging within 3-4 hours
User-friendly:  Comfortable grip 

Testimonial: “It works perfectly and for half the price as in normal stores, where they don't even come with accessories.” - Justin

Introducing the Sofrid Vacuum Pro, the ultimate cleaning companion that takes your cleaning routine to the next level. Equipped with a powerful cyclone suction system and a robust 120W motor, this vacuum effortlessly tackles even the most stubborn hair, debris, and microscopic particles that lurk in your living spaces. Its versatility knows no bounds, thanks to the inclusion of various attachments like the extended hose, crevice tool, and dust brush, allowing you to effortlessly reach every nook and cranny, delicate surfaces, and intricate interior decorations with ease.

What Makes It So Special?
The Sofrid Vacuum Pro is a truly exceptional cleaning companion that stands out from the rest. Its powerful 120W motor and cyclone suction technology effortlessly tackle hair, debris, and microscopic particles, while the included attachments provide versatile cleaning options for hard-to-reach areas. The cordless and lightweight design offers unparalleled maneuverability, and the fast charging 3500mAh li-ion batteries ensure a 30-minute runtime for uninterrupted cleaning. With intelligent power management technology and multiple safety features, the Sofrid Vacuum Pro guarantees a worry-free cleaning experience. Say goodbye to mediocre cleaning and embrace the outstanding performance and convenience of the Sofrid Vacuum Pro.

A Customized Cleaning Experience 
The Sofrid Vacuum Pro is highly adaptable and customizable with its range of included accessories. Whether you need to clean narrow corners, delicate surfaces, or interior decorations, this vacuum has you covered. The extended hose widens the cleaning scope, making it easier to reach tight spaces. The crevice tool is perfect for picking up debris in narrow corners, while the dust brush efficiently removes hair and residue from delicate surfaces. With the adjustable accessories, you can easily tailor the vacuum's cleaning capabilities to suit your specific needs, ensuring a thorough and precise cleaning experience every time.

Features
Powerful Cyclone Suction: The Sofrid Vacuum Pro is equipped with a powerful cyclone suction system that effortlessly captures hair, debris, and microscopic particles, ensuring a thorough and deep clean.
Versatile Attachments: This vacuum comes with a range of attachments including an extended hose, crevice tool, and dust brush. These accessories allow you to clean hard-to-reach areas, narrow corners, and delicate surfaces with ease.
Cordless and Lightweight: Enjoy the freedom of cordless cleaning with the Sofrid Vacuum Pro. Its lightweight design enables easy maneuverability and access to every corner of your home, without the hassle of cords or heavy equipment.
Fast Charging and Long Battery Life: The vacuum features a quick charging time of 3-4 hours, providing you with ample cleaning time. The long-lasting battery ensures uninterrupted cleaning sessions, allowing you to tackle your cleaning tasks efficiently.

Trusted and Recommended by Experts
The Sofrid Vacuum Pro has gained the trust and recommendation of experts in the field of cleaning appliances. Renowned cleaning professionals and experts have recognized its exceptional performance and advanced features. With its powerful suction, versatile attachments, and user-friendly design, the Sofrid Vacuum Pro has become a go-to choice for those seeking reliable and effective cleaning solutions. Its quality construction and attention to detail have earned it a solid reputation among experts, making it a trusted and recommended option for achieving exceptional cleaning results.

Fast Charge, Long Runtime
The Sofrid Vacuum Pro features fast charging technology that allows you to quickly recharge the vacuum's batteries in just 3-4 hours. This means that you can spend less time waiting for the vacuum to charge and more time cleaning. With its efficient charging capability, you can enjoy uninterrupted cleaning sessions without worrying about running out of battery power. The fast charging feature adds convenience and efficiency to your cleaning routine, ensuring that the vacuum is always ready to tackle your cleaning tasks whenever you need it.

Durable and User-Friendly Design
The Sofrid Vacuum Pro boasts a durable and user-friendly design that sets it apart from other vacuum cleaners on the market. Built with high-quality materials, it is engineered to withstand the rigors of regular use and deliver long-lasting performance. Its sturdy construction ensures that it can withstand the demands of various cleaning tasks without compromising its functionality. Additionally, the user-friendly design enhances the overall cleaning experience. From the ergonomic handle that provides a comfortable grip to the intuitive controls that make operation effortless, every aspect of the design is geared toward ease of use. Whether you're a seasoned cleaner or a first-time user, the thoughtful design of the Sofrid Vacuum Pro ensures that anyone can operate it with confidence and achieve excellent cleaning results.

Satisfaction Guarantee

14 DAYS MONEY BACK GUARANTEE
We'll refund your money if you're not 100% satisfied!
Order now with confidence! If for any reason you don't think Sofrid Vacuum Pro is for you, we offer a 14 day money-back guarantee. So if you don't love it, you can get your money back. No questions asked!


"""})

    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=messages
    )

    messages.clear()
    return response.choices[0].message.content

def generate_blogs(data):
    # Check if the folder exists. If not, create one.
    folder_name = "Generated_Blogs"
    if not os.path.exists(folder_name):
        os.mkdir(folder_name)

    titles = [doc['Title'] for doc in data]

    progress_bar = tqdm(titles, desc="Generating Blogs", unit="blog", bar_format="{desc}: {percentage:3.0f}%|{bar}| {n_fmt}/{total_fmt}", colour='green', ncols=80)

    for title in progress_bar:
        document = Document()  # Create a new document for each blog
        blog_content = generate_blog_content(title)
        document.add_heading(title, level=1)
        sentences = blog_content.split('. ')
        paragraph_text = '. '.join(sentences)
        document.add_paragraph(paragraph_text)
        document.save(f"{folder_name}/{title}.docx")  # Save the blog in the folder

        progress_bar.set_postfix({"Current Title": title})
        progress_bar.update(1)

def main():
    # Use the file finder function to get the file path
    print("Please select your CSV file...")
    file_path = select_file_path()

    # If no file is selected, exit the program
    if not file_path:
        print(Fore.RED + "No file selected. Exiting program." + Style.RESET_ALL)
        return

    data = load_data_from_csv(file_path)

    

    generate_blogs(data)

    clear_screen()
    print(Fore.LIGHTGREEN_EX + "Blog generation completed." + Style.RESET_ALL)

if __name__ == '__main__':
    main()
