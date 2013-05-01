<%

dim config_site_name,	config_site_url
dim config_site_url_ssl
dim config_error_name,	config_error_email
dim config_default_template
dim config_cookie_domain, config_cookie_url, config_cookie_expire
dim config_cache_dir
dim config_db

config_site_name 		= "My Sample Website"
config_site_url 		= "http://localhost/"
config_site_url_ssl 	= "https://localhost/"

config_error_name 		= "*Simplicity ASP Framework"
config_error_email 		= "website_errors@domain.com"

config_default_template	= "template.asp"

config_cookie_domain 	= "simpleCore"
config_cookie_url		= "localhost"
config_cookie_expire	= date+30

config_db 				= "dev"

' Make sure this dir has write access with IUSR_MACHINE user
config_cache_dir		= "C:\hosts\framework\system\cache\"

' ---------------------------------------------------------------
'  Setting Website Defualts 
' ---------------------------------------------------------------
' 
' these defaults can be overwritten at anytime
' within this application
' 

dim default_site_name, default_template

default_site_name	= config_site_name
default_template 	= config_default_template

%>