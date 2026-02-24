import re

def clean_wix_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # Patterns to remove Wix-related code
    patterns = [
        r'<meta name="generator" content="Wix\.com Website Builder">',  # Wix generator meta tag
        r'<meta http-equiv="X-Wix-Meta-Site-Id".*?>',  # Wix meta site ID
        r'<meta http-equiv="X-Wix-Application-Instance-Id".*?>',  # Wix application instance ID
        r'<meta http-equiv="X-Wix-Published-Version".*?>',  # Wix published version
        r'<script.*?src=".*?wix.*?".*?>.*?</script>',  # Wix-related script tags
        r'<style.*?data-url=".*?wix.*?".*?>.*?</style>',  # Wix-related style blocks
        r'<!--pageHtmlEmbeds.head start-->.*?<!--pageHtmlEmbeds.head end-->',  # Wix head embeds
        r'<!--pageHtmlEmbeds.bodyStart start-->.*?<!--pageHtmlEmbeds.bodyStart end-->',  # Wix body embeds
        r'<style.*?data-href=".*?wix.*?".*?>.*?</style>',  # Wix Thunderbolt styles
    ]

    # Remove all matching patterns
    for pattern in patterns:
        content = re.sub(pattern, '', content, flags=re.DOTALL)

    # Write the cleaned content back to the file
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(content)

    print(f"Cleaned Wix-related code from {file_path}")

# Specify the HTML file to clean
html_file_path = r"c:\temp\PowerShell(using)\Home _ Oss Powershell Scrip(2).html"
clean_wix_html(html_file_path)