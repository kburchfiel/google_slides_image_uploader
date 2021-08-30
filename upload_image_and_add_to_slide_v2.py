# Script for Adding Images Stored on Computer Into Google Slides 
# By Ken Burchfiel
# Released under the MIT License
# (This code makes extensive use of various Google code excerpts. I believe these snippets all use the Apache 2.0 license.)

# For additional background on this program, see the slide_update_script.ipynb file, which applies this code to add images to a sample presentation.


# This file contains various functions that allow images to be inserted into Google Slides presentations. These functions include:
# 
# 1. upload_blob(), which uploads an image to Google Cloud Storage
# 2. generate_download_signed_url_v4(), which creates a signed URL that can be utilized for the image import
# 3. add_image_to_slide(), which uses the URL created in generate_download_slide_url to import the image in Google Cloud Storage into Google Drive
# 4. delete_blob(), which removes the image from the Cloud Storage Bucket
# 5. upload_image_and_add_to_slide, which applies the above four functions to complete the image import process. This function is then called by slide_update_script.ipynb for each image/slide pair specified in that script.



# First, I'll import a number of libraries and modules that will be relevant for the code.

import time

# The following two import lines come from the "Creating a Signed URL to upload an object" code sample available at https://cloud.google.com/storage/docs/access-control/signing-urls-with-helpers#code-samples_1 . This sample uses the Apache 2.0 license.
import datetime
from google.cloud import storage

# The following import statements come from the Slides API Python quickstart (https://developers.google.com/slides/api/quickstart/python), which uses the Apache 2.0 license.

import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from google.oauth2 import service_account # From https://developers.google.com/identity/protocols/oauth2/service-account#authorizingrequests , which uses the Apache 2.0 license.


def upload_blob(credentials, bucket_name, source_file_path, destination_blob_name): # The original source of this code, comments, and docstring was https://cloud.google.com/storage/docs/uploading-objects#storage-upload-object-code-sample ; the code was made available under the Apache 2.0 license. I made some minor modifications to the code.
    """Uploads a file to the bucket."""

    # The ID of your GCS bucket
    # bucket_name = "your-bucket-name"
    # The path to your file to upload
    # source_file_name = "local/path/to/file"
    # The ID of your GCS object
    # destination_blob_name = "storage-object-name"

    storage_client = storage.Client(credentials=credentials) # The default argument for credentials is None--see https://googleapis.dev/python/storage/latest/client.html . I edited this line to use the credentials passed in from upload_image_and_add_to_slide().
    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(destination_blob_name)

    blob.upload_from_filename(source_file_path)
    # To avoid a 403 error with the above line, I also needed to add the storage admin role to my service account within the IAM settings in the Google Cloud Console. See https://stackoverflow.com/a/56305946/13097194

    print("File uploaded to cloud.")

def generate_download_signed_url_v4(bucket_name, blob_name, credentials): 
    # The original source of this code, comments, and docstring was https://cloud.google.com/storage/docs/access-control/signing-urls-with-helpers#storage-signed-url-object-python ; the code was made available under the Apache 2.0 license. I made some minor modifications to the code.

    """Generates a v4 signed URL for downloading a blob.

    Note that this method requires a service account key file. You can not use
    this if you are using Application Default Credentials from Google Compute
    Engine or from the Google Cloud SDK.
    """
    # bucket_name = 'your-bucket-name'
    # blob_name = 'your-object-name'

    storage_client = storage.Client(credentials=credentials)
    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(blob_name)

    url = blob.generate_signed_url(
        version="v4",
        expiration=datetime.timedelta(minutes=1),
        # I made this URL valid for 1 minute only (see below) since it will be used immediately and will only be needed once. 
        # Allow GET requests using this URL.
        method="GET",
    )

    print("Signed URL generated.")
    return url

# The following function (add_image_to_slide()) uses the image URL returned from generate_download_signed_url_v4 to attach an image to a slide within a Google Slides presentation. 

def add_image_to_slide(presentation_id, credentials, image_url, image_name, page_object_id, scaleX = 1.8, scaleY = 1.8, translateX = 1100000, translateY = -500000): 
    # The scaleX, scaleY, translateX, and translateY values align each image with its corresponding slide. The translateX and translateY values are in EMU (English Metric Units), so you can use the 914,400 EMUs per inch conversion as a reference when appending slides. I recommend testing the scale and translation values on one image, then expanding the for loop to include all images once you have the positioning right.
    # To find the presentation ID, look for the string of characters that proceeds the /d/ section of the URL for a Google Slides document. For instance, if a presentation's URL is https://docs.google.com/presentation/d/1xJfItB6w7hH0Nq2va-B1nUJFKGto_sFKxdMRvrMRvsI/ , the last part of that URL (1xJfItB6w7hH0Nq2va-B1nUJFKGto_sFKxdMRvrMRvsI), excluding the forward slash, is the presentation's ID. 

    service = build('slides', 'v1', credentials=credentials)
    # Call the Slides API. From the Slides API Python Quickstart (https://developers.google.com/slides/api/quickstart/python); the Quickstart code sample uses the Apache 2.0 license.

    # The following code derives from the 'Adding Images to a Slide' Python code snippet available at https://developers.google.com/slides/api/guides/add-image . This code uses the Apache 2.0 license.

    emu4M = {
        'magnitude': 4000000,
        'unit': 'EMU'
    }

    requests = []

    # The following set of code checks whether the image is already present on the slide. It does so by generating a list of all object IDs and checking whether the name of the image is present within those IDs.
    page_info = service.presentations().pages().get(presentationId = presentation_id, pageObjectId = page_object_id).execute() # Retrieves a wealth of information about the current slide. This line is adapted from https://googleapis.github.io/google-api-python-client/docs/dyn/slides_v1.presentations.pages.html#get ; I believe the code on that page also uses the Apache 2.0 license.
    objects_on_page = []
    # print(page_info) # For debugging
    
    # The documentation for the get() function at https://googleapis.github.io/google-api-python-client/docs/dyn/slides_v1.presentations.pages.html#get will serve as a helpful reference for understanding and modifying this code.

    if 'pageElements' in page_info.keys(): # This check was included so that the following code would be skipped if there weren't any page elements on the slide. Without this check, the program would return a KeyError for blank slides due to the missing pageElements key.
        #print('pageElements found in keys')
        for k in range(len(page_info['pageElements'])): # i.e. for each separate element on the page. Each member of the pageElements list is a dictionary, and objectId (accessed below) is a key that retrieves the object id within that dictionary.
            objects_on_page.append(page_info['pageElements'][k]['objectId']) # 

        if image_name in objects_on_page: # Checks whether the image planned for addition to the slide is already present. If so, the following line files a request to delete that image. This assumes that both the image on the slide and the image planned to be added to the slide have the same name.
            requests.append({'deleteObject':{'objectId':image_name}}) # This code is based on the batchUpdate() function available at https://developers.google.com/resources/api-libraries/documentation/slides/v1/python/latest/slides_v1.presentations.html. The code uses the Apache 2.0 license unless I'm mistaken. See https://developers.google.com/open-source/devplat
            print("Image already present on slide. Request added to delete the pre-existing copy so that a new one can be added.")
    #else:
    #    print('pageElements not found in keys') # For debugging

# The following code, which provides information about a new image to be added to the presentation, also comes from the 'Adding Images to a Slide' code sample, though I made some minor modifications.

    requests.append({
        'createImage': {
            'objectId': image_name,
            'url': image_url,
            'elementProperties': {
                'pageObjectId': page_object_id,
                'size': {
                    'height': emu4M,
                    'width': emu4M
                },
                'transform': {
                    'scaleX': scaleX,
                    'scaleY': scaleY,
                    'translateX': translateX,
                    'translateY': translateY,
                    'unit': 'EMU'
                }
            }
        }
    })

    body = {
        'requests': requests
    }
    response = service.presentations() \
        .batchUpdate(presentationId=presentation_id, body=body).execute() # Based on both the batchUpdate function and the Slides API Python Quickstart. I originally  # Originally started with response = slides.service but this Stack Overflow answer by Shubham Kushwaha (https://stackoverflow.com/a/59096196/13097194) indicated that I could just use service instead of slides_service.


    print("Image added to slide. Pre-existing image deleted if requested.")

def delete_blob(bucket_name, blob_name, credentials): # Source of original code: https://cloud.google.com/storage/docs/deleting-objects#code-samples Code uses the Apache 2.0 License .
    """Deletes a blob from the bucket."""
    # bucket_name = "your-bucket-name"
    # blob_name = "your-object-name"

    storage_client = storage.Client(credentials=credentials) # Updated this line to use credentials from the upload_image_and_add_to_slide function

    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(blob_name)
    blob.delete()

    print("Blob {} deleted.".format(blob_name))


def upload_image_and_add_to_slide(image_folder_path, image_file_name, image_file_extension, service_account_path, scopes, presentation_id, page_object_id, bucket_name, scaleX, scaleY, translateX,translateY):
    
    '''
    This function performs 5 steps:
    Step 1: Uploads an image to Google Cloud Storage (via the upload_blob function)

    Step 2: Creates a signed URL that can be used to copy the uploaded image into Google Slides (via the generate_download_signed_url_v4 function)

    Step 3: Deletes the previous copy of the image from Google Slides (if one exists) using the add_image_to_slide() function

    Step 4: Copies the image into Google Slides via the add_image_to_slide() function

    Step 5: Deletes the image from Google Cloud Storage using the delete_blob() function

    Variable descriptions:
    image_folder_path: The path to the folder containing the image, which shouldn't be confused with the path to the image itself
    image_file_name: the name of the image within that folder. If you're adding images to a slide within a for loop, consider storing the image file names within a list and iterating through that list. Alternately, you could store both the image file names and page object ids as values within a list of lists.
    image_file_extension: .png, .jpg, etc. Separated from image_file_name because I don't think Google allows object IDs to have periods
    service_account_path: the path to the credentials.json account stored on your computer
    page_object_id: the id of the slide being updated by the function. Needs to be retrieved within Google
    bucket_name: the name of the Google Cloud Storage bucket containing the image file
    scaleX, scaleY, transformX, and transformY can be adjusted to set the correct position and size of the image on the slide.
    '''

    # The following two lines of code derive from the 'Preparing to make an authorized API call' code samples available at https://developers.google.com/identity/protocols/oauth2/service-account#authorizingrequests . These code samples use the Apache 2.0 license. 
    SERVICE_ACCOUNT_FILE = service_account_path 
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)

    image_file_path = image_folder_path + image_file_name + image_file_extension # i.e. ..\\pictures\image_1.png'
    
    # Step 1
    upload_blob(credentials,bucket_name,image_file_path, destination_blob_name = image_file_name)
    
    # Step 2
    url = generate_download_signed_url_v4(bucket_name, blob_name = image_file_name, credentials = credentials)
    time_image_becomes_accessible = time.time()
    
    # Steps 3 and 4
    add_image_to_slide(presentation_id = presentation_id, credentials = credentials, image_url = url, page_object_id=page_object_id, image_name = image_file_name, scaleX=scaleX, scaleY=scaleY, translateX=translateX,translateY=translateY)

    # Step 5
    delete_blob(bucket_name = bucket_name, blob_name = image_file_name, credentials = credentials)
    time_image_is_no_longer_accessible = time.time()
    time_image_was_accessible = time_image_is_no_longer_accessible - time_image_becomes_accessible
    print("Image was accessible for",'{:.4f}'.format(time_image_was_accessible),"second(s).")
    # On my computer, the image was accessible for 2.10-2.37 seconds (based on my first four tests); however, the image can be accessible for longer periods in some cases (but not longer than 1 minute, the length that the signed URL is valid).
