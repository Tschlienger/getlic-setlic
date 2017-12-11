# getlic-setlic
Simple Windows license export and activation tool.

## How to use
Use getlic.ps1 to save the current product key to a license file.
Then, use setlic.ps1 to search for keys in the license file matching your current Windows version/edition and activate the os.

## Usage example
I'm setting up Microsoft Surface devices for my company and I use images to do so quickly. Unfortunately, cloning devices causes the pre-installed license keys to get lost.
With the getlic.ps1 script on a USB stick, I can simply export the key from a new Surface, clone it, and install our KMS key instead.
The keys are stored on my USB stick as well and whenever I got a new PC or Laptop, I can simply run the setlic.ps1 script to activate it with one of my collected licenses.
