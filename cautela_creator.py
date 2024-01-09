import os, time
from cautela_other_functions import replace_words_from_dictionary, create_item_table_paragraph_style, open_document
from cautela_input_functions import get_non_repeatable_user_inputs, get_item_table_user_inputs

if __name__ == "__main__":

    username = os.getlogin()
    # Looks for the file "Cautela Template" in the desktop of the current 
    # user, and exits the program if the file is not found
    document = open_document(username)

    # Creates a dictionary of the one time user inputs, then replaces each word
    # in the template with the inputs.
    non_repeatable_inputs = get_non_repeatable_user_inputs()
    replace_words_from_dictionary(document, non_repeatable_inputs)

    # Creates the style used by items in the item table, 'Table Item'
    create_item_table_paragraph_style(document)

    # Starts a loop to collect informations about each item to be added to 
    # the list until the user exits it
    get_item_table_user_inputs(document)

    # Follows the naming convention of the current template
    cautela_name = "{}.{} - {} - Envio de Equipamentos de {} para {}.docx".format(
        non_repeatable_inputs["CAUTELAID"], 
        non_repeatable_inputs["YEAR"], 
        non_repeatable_inputs["GLPINUMBER"], 
        non_repeatable_inputs["SENDERUNIT"], 
        non_repeatable_inputs["RECEIVERUNIT"]
    )
    print("Cautela salva como: {}".format(cautela_name))
    document.save(f"C:/Users/{username}/Desktop/" + cautela_name)
    time.sleep(5)
