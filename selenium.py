for item in selected_data:
    cnpj, senha = item

    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)

    navegador.get("https://docs.google.com/forms/u/1/d/e/1FAIpQLScKrG3u9_CCAuO6AiJ74luw9AFdz8Ow56KKPm_eL7T0o0VGaQ/formResponse")
    navegador.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/c-wiz/main/div[2]/div/div/div[1]/form/span/section/div/div/div[1]/div[1]/div[1]/div').send_keys(email)
    navegador.find_element(By.XPATH, '//*[@id="identifierNext"]').click()

    
