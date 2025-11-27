import input_handler
import config_reader


config_reader = config_reader.ConfigReader("config.json")

def main():
    input_handler_maker_name:str = config_reader.get("input_handler_name",
                                                     "分组多输出输入器")
    input_handler_maker: type[input_handler.ABC_输入器] = getattr(input_handler_maker_name,
        input_handler_maker_name, input_handler.分组多输出输入器)
    a_input_handler = input_handler_maker(config_reader=config_reader)
    a_input_handler.main()
    # input("程序结束")

if __name__ == "__main__":
    main()