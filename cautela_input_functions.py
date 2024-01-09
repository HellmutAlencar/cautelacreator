import datetime, time, os
from cautela_other_functions import add_item_to_item_table

# Used to set the max length of the cautela id, since in no year the ID reached 4 digits in length
MAX_CAUTELAID_STRING_LENGTH = 3

# Used to set the size of a GLPI id, which at the moment is a 5-digit number
GLPI_ID_STRING_LENGTH = 5

def get_non_repeatable_user_inputs():

    year = str(datetime.date.today().year)
    # Dictionary to list the most frequent destinations, making it faster to fill the form in typical cases
    dictionary_of_units = {
        "1" : "DTIC-SEDE",
        "2" : "DTIC-PB",
        "3" : "DTIC-AL",
        "4" : "DTIC-BH",
    }

    cautela_id = input("Digite o número da cautela (ex:112): ")
    while True:
        if not cautela_id.isnumeric(): # Cautela ID must be a number
            print("O ID da cautela deve ser um número (ex:145)")
            cautela_id = input("Digite o número da cautela (ex:112): ")

        elif len(cautela_id) > MAX_CAUTELAID_STRING_LENGTH: # Cautela ID must not exceeded 3 characters
            print("O número inserido é grande demais (max:3 dígitos)")
            cautela_id = input("Digite o número da cautela (ex:112): ")

        else:
            break

    glpi_number = input("Digite o número do GLPI (ex:64567): ")
    while True:
        if not glpi_number.isnumeric(): # GLPI must be a number
            print("O ID do GLPI deve ser um número (ex:65732)")
            glpi_number = input("Digite o número do GLPI (ex:64567): ")

        elif len(glpi_number) != GLPI_ID_STRING_LENGTH: # GLPI must have 5 digits
            print("O número inserido deve ter 5 dígitos")
            glpi_number = input("Digite o número do GLPI (ex:64567): ")

        else:
            break

    sender_unit = input("Digite o nome da unidade de ORIGEM (ex:DTIC-SEDE): \n1 - DTIC-SEDE\n2 - DTIC-PB\n3 - DTIC-AL\n4 - DTIC-BH\n5 - Outro\nSua escolha: ")
    while True:
        if sender_unit in ('1', '2', '3', '4'):
            sender_unit = dictionary_of_units[sender_unit]
            break

        elif sender_unit == '5':
            sender_unit = input("Digite o nome da unidade de ORIGEM (ex:IJI-DELEG): ").upper()
            break

        else:
            sender_unit = input("Opção inválida, tente novamente: ")

    receiver_unit = input("Digite o nome da unidade de DESTINO (ex:DTIC-AL): \n1 - DTIC-SEDE\n2 - DTIC-PB\n3 - DTIC-AL\n4 - DTIC-BH\n5 - Outro\nSua escolha: ")
    while True:
        if receiver_unit in ('1', '2', '3', '4'):
            receiver_unit = dictionary_of_units[receiver_unit]
            break

        elif receiver_unit == '5':
            receiver_unit = input("Digite o nome da unidade de DESTINO (ex:IJI-DELEG): ")
            break

        else:
            receiver_unit = input("Opção inválida, tente novamente: ")

    receiver_room = input("Digite qual setor vai receber (ex:43PROM): ").upper()

    username = os.getlogin()
    intern_dictionary = {}
    try: # Will succeed if the .txt is valid
        with open(f"C:/Users/{username}/Desktop/estagiários.txt") as f:
            for line in f:
                (key, val) = line.split()
                intern_dictionary[str(key)] = val.replace("_", " ")
    except FileNotFoundError: # File missing or wrong file name
        print("ERRO\nFalta o arquivo estagiários.txt com esse exato nome na mesma "
              "pasta do programa e seu conteúdo deve ter a seguinte estrutura, com _ nos espaços dos nomes:"
              "\n1 - Michael_Jackson\n2 - Antonio_Carlos_Jobim")
        time.sleep(20)
        exit()
    except ValueError: # Missing an underline or wrong file structure
        print("ERRO\nÉ necessário que o arquivo estagiários.txt na mesma "
              "pasta do programa tenha a seguinte estrutura, com _ nos espaços dos nomes:"
              "\n1 - Michael_Jackson\n2 - Antonio_Carlos_Jobim")
        time.sleep(20)
        exit()
    except: # Exception for any other error
        print("ERRO\nAlgo deu errado.")
        exit()

    print("Digite o nome de quem vai enviar:")
    for intern_index, intern_name in intern_dictionary.items():
        print("{} - {}".format(intern_index, intern_name))
    sender = (input("0 - Outro\nSua escolha: "))

    while(True):
        if sender == '0':
            sender = input("Digite o nome da pessoa que vai enviar: ").title()
            break
        try:
            sender = intern_dictionary[sender]
            break
        except:
            sender = input("Opção inválida, tente novamente: ")

    dictionary = {
        "CAUTELAID" : cautela_id,
        "YEAR" : year,
        "RECEIVERUNIT" : receiver_unit,
        "RECEIVERROOM" : receiver_room,
        "GLPINUMBER" : glpi_number,
        "SENDERUNIT" : sender_unit,
        "SENDERNAME" : sender
    }

    return dictionary

def get_item_table_user_inputs(document):
    
    item_number = 1
    item_name = input("Digite a descrição do equipamento (ex:Computador Lenovo M80q): ").title()
    item_code = input("Digite o tombo do equipamento (ex:021342): ")
    item_state = input("Digite o estado do equipamento (ex:Bom): ").title()

    item_information = (
        str(item_number),
        item_name,
        item_code,
        item_state
    )

    add_item_to_item_table(document, item_information)

    while True:
        collect_input = input("Digite a descrição do equipamento (ex:Computador Lenovo M80q): \nOu digite 1 para que a descrição seja igual ao item anterior: \nOu aperte ENTER para terminar: ").title()
        if collect_input == '':
            return
        elif len(collect_input) != 1:
            item_name = collect_input

        collect_input = input("Digite o tombo do equipamento (ex:021342): ").upper()
        if len(collect_input) != 1:
            item_code = collect_input

        collect_input = input("Digite o estado do equipamento (ex:Bom): \nOu digite 1 para que a descrição seja igual ao item anterior: ").title()
        if len(collect_input) != 1:
            item_state = collect_input

        item_number += 1
        item_information = (
            str(item_number),
            item_name,
            item_code,
            item_state
        )

        add_item_to_item_table(document, item_information)
