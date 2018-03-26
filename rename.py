import os

no_of_chars_to_remove = 28  # 24 for International
path_dir = 'India'  # Intl for international
path = 'config_files/' + path_dir
os.chdir(path)

for filename in os.listdir('.'):
    os.rename(filename, filename[:len(filename)-no_of_chars_to_remove]+'.txt')

