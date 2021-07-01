import time
# set the parameters of the serial
port = "COM7"
baudrate = 115200


# commands will be used
exiT = "AT+EXIT\r\n"
tri_plus = "+++\r\n"
reset = "AT+DEF\r\n"
request1 = "AT+SYMBOL=?\r\n"
request2 = "AT+FREQ?\r\n"
request3 = "AT+MODEM=?\r\n"

# the list of all the commands will be called in this test
command_seq = [exiT, request1, reset, request2, tri_plus, request3]

# the head of the excel form
key_list = ["指令", "返回", "时间"]

# the path to store the result
path_excel = "./result_{}.xls".format(int(time.time()))

# error code
error1 = "Command not support yet!"
error2 = "ERROR,1"
