#importing Image class from PIL package
from PIL import Image

#import requests to get content from metmuseum api
#import json to parse responses from metmuseup api
#import random for selecting an object at random from the chosen department
#import os to allow for filename to be selected at random
#import the time library to allow sleep
import requests, json, random, os, time
#import BytesIO to parse the response
from io import BytesIO

from datetime import datetime

import openpyxl # to read and write xls files

#@Excel functions to write to excel
# create an excell sheet using openpyxml and write it to file
# returns a workbook object that can be saved to file
def create_xls_from_list(_list, wb):
    # the the workbook active sheet
    sheet = wb.active
    sheet.append(_list)
    return wb


# save a generated file to the file system
def save_workbook(wb, filename):
     wb.save(filename)

#@utility function to create a dict object from a list of dicts
#@params taken are a list of dicts and the key to be removed
#here in our case we only want to remove the name 'departmentId'
#this is because the number will correspond with the department
#this number can then be used to select a department
def create_dict(list_of_dicts, removed_key):
    new_dict = {}
    if list_of_dicts is not {}:
        for item in list_of_dicts:
            name = item.pop(removed_key)
            new_dict[name] = item
    else:
        new_dict['Error'] = "The dict provided was empty"
    return new_dict


#metmuseum api handling
#using requests library
#@function to return a list of departments for the user
def get_department_names(base_url):
    #append the route to the departments to the url and send a get request
    try:
        #check if the request worked
        response = requests.get(base_url+"/departments")
        #save the result in a variable after converting it to json
        #only save the results under the departments key
        departments_object =list(json.loads(response.content)['departments'])
        #use list comprehension to save the department names in a list
        # department_names = {for x in departments_object}
        #return that list
        return create_dict(departments_object, 'departmentId')
    except Exception as e:
        #return an error message
        print(e)
        error_message = "Oops, looks like there was a problem!"
        return error_message


#here we define a function to fetch specific details about an object
#title of the art image : @param title
#artist of the art image :  @param artistDisplayName
#url of the full sized art Image : @param primaryImage
#url of the small sized art Image : @param primaryImageSmall
#bool of the Image and if it exists in the public domain or #not
#@param : isPublicDomain this is used to give feedback on the program
def get_art_object(base_url, object_id):
    object_url = '/objects/'+str(object_id)
    try:
        response = requests.get(base_url+object_url)
        object_information = json.loads(response.content)
        if object_information['isPublicDomain'] == True:
            #here a dict is created to hold necessary params needed from the
            #whole object
            selected_object = {}
            selected_object['title'] = object_information['title']
            selected_object['artistDisplayName'] = object_information['artistDisplayName']
            selected_object['primaryImage'] = object_information['primaryImage']
            selected_object['primaryImageSmall'] = object_information['primaryImageSmall']
            for x in selected_object.keys():
                if selected_object[x] == '':
                    selected_object[x] = 'Unknown'
            return selected_object
        else:
            return "Error: The selected art piece is not in public domain"

    except Exception as e:
        #print the error
        print(e)
        error_message = "Oops, looks like there was a problem!"
        return error_message


#here we get the objects in a particular department using the departmentID
#@params are the departmentId and the base url
def random_department_object(base_url, dept_id):
    object_url = base_url+"/objects?departmentIds="+str(dept_id)
    try:
        response = requests.get(object_url)
        object_ids = json.loads(response.content)['objectIDs']
        #create a random number inorder to select a random objectID
        random_object_id = random.randint(0,len(object_ids))
        return object_ids[random_object_id]
    except Exception as e:
        #print the error_message
        print(e)
        error_message = "Oops, looks like there was a problem!"
        return error_message

#@Utitlity function to split a url into the function
def split_url(url):
    return url.split("/")[-1]



#@function that uses pillow to get an image
#given an image url
def image_from_url(image_url):
    try:
        #use requests to get image response from url
        response =  requests.get(image_url)
        #display image by rendering it using BytesIO
        image = Image.open(BytesIO(response.content))
        #return the image for any manipulation
        return image
    except Exception as e:
        print(e)
        return "Sorry, the image was not retrieved"

#this @function is used to select a random sticker from the images folder
def select_random_sticker():
    #initialize the sticker sticker_directory
    #check if it exists and if not create it
    #then recurse
    try:
        os.makedirs('saved_images')
        select_random_sticker()
    except OSError as e:
        sticker_directory = 'stickers/'
        #choose a random sticker and initialize a PIL Image file
        random_sticker_file = random.choice(os.listdir(sticker_directory))
        random_sticker = Image.open(sticker_directory+random_sticker_file)
        sticker_filename = random_sticker_file.split('.')[0]
        #convert the image to RGBA to allow it to be overlayed on the result image
        return (random_sticker.convert('RGBA'), sticker_filename)


#this @function takes dimensions and returns a random tuple
#that is less that those dimensions provided
def random_dims(dim_x, dim_y):
    #divide the first number by 2 to allow the sticker to be in view
    new_dim_x = random.randint(0, dim_x//2)
    #divide the second number by 2 to allow the sticker to be in view
    new_dim_y = random.randint(0, dim_y//2)
    #return the tuple
    return (new_dim_x, new_dim_y)

#this function saves an image given a sticker name and an image
#@param : image, filename
#the functions returns success or fail
def save_stickered_image(image, filename):
    try:
        try:
            os.makedirs('saved_images')
        except OSError as e:
            image.save('saved_images/'+filename)
            #return a success error_message
            return "Image saved Successfully"
    except Exception as e:
        print(e)
        #return an error message
        return "Sorry, the image couldn't be saved"


#this @function is used to apply a sticker to the artwork
#the sticker is randomly rotated and or flipped
# and placed in a random location on the image
#takes an object from the Image class
def apply_sticker(image, sticker):
    #flip or rotate sticker to a random degree
    #apply sticker at a random point
    #get the size of the original image object
    x,y = image.size

    #flip the sticker before applying
    sticker = sticker.rotate(random.randint(10,300))
    #paste the sticker at a random point in the image with the sticker overlayed
    #last param is a mask
    image.paste(sticker,random_dims(x,y), sticker)
    #return the image object
    return image

#here be dragons
def main():
    #introductory text
    print("Hello Welcome to the Met Museum Program")
    #provide the met museum  base_url
    #define the base url for metmuseum's api
    museum_base_url = "https://collectionapi.metmuseum.org/public/collection/v1"
    #Loop through to allow the program to run continiously
    while 1:
        #Show a list of departments and get a user to select one
        list_of_departments = get_department_names(museum_base_url)
        for x,y in list_of_departments.items():
            print(str(x) +":"+ y['displayName'])
        selected_department = int(input("Please select a department: "))
        department_object= random_department_object(museum_base_url, selected_department)
        object_dict = get_art_object(museum_base_url,department_object)
        if(type(object_dict) != type('')):
            selected_image = image_from_url(object_dict['primaryImageSmall'])
            selected_image.show()
            chose_image = input("Do you like the selected Image (Y/N)")
            if chose_image == 'Y':
                #Apply a random sticker to the image and save it as well as show it
                sticker, sticker_filename = select_random_sticker()
                apply_sticker(selected_image, sticker)
                selected_image.show()
                #here we get the original image filename from a url and get the filename
                #of the original image as the last part of the url
                original_image_filename = str(split_url(object_dict['primaryImage'])).split('.')[0]
                filename_to_save = original_image_filename+sticker_filename+'.jpg'
                #prepare dict for write to excel
                object_dict['sticker'] = sticker_filename+'.png'
                object_dict['primaryImage'] = split_url(object_dict['primaryImage'])
                del object_dict['primaryImageSmall']
                #get current time
                current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
                object_dict['timestamp'] = current_time
                save_stickered_image(selected_image, filename_to_save)
                #create headers for the excel sheet
                excel_headers = ['Title', 'Artist Display Name', 'Original Art Image', 'Sticker File', 'Timestamp']
                print(list(object_dict.values()))
                #create a workbook
                wb = openpyxl.Workbook()
                create_xls_from_list(excel_headers, wb)
                create_xls_from_list(list(object_dict.values()), wb)
                save_workbook(wb, 'metmuseum.xls')
        else:
            print("Sorry, we couldn't find any image in the public domain")
            time.sleep(1)
            print("Please try again")
            time.sleep(1)
            main()



#    current_time = datetime.now().strftime("%Y-%m-%d %H:%M")

if __name__ == '__main__':
    main()