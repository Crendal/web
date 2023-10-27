# 서버 띄우기   flask 는 기본으로 port를 5000으로 하는구나
from flask import Flask
import argparse
app = Flask(__name__)
# print('TEST2')
# 아무것도  path에 안 붙어 있는 경우의 처리

parser = argparse.ArgumentParser()


@app.route("/",methods=['GET','POST'])
         # 경로,    방식
def main():
    return "Hello"


# 인자값을 받을 수 있는 인스턴스 생성
parser = argparse.ArgumentParser(description='사용법 테스트입니다.')

# 입력받을 인자값 등록
parser.add_argument('--target', required=True, help=' write your port')


# 입력받은 인자값을 args에 저장 (type: namespace)
args = parser.parse_args()

# 입력받은 인자값 출력
print(args.target)
print(args.env)

# 중개함수 (경로 방식)

# main 으로 호출 되었을 때만 실행됨
    # 다른 거에서 import해서 사용하면 실행 안됨 
if __name__=='__main__':
# print('TEST1')    
    app.run(debug=True, port= args.port)

# 이렇게 하면 not found


