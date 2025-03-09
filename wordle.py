import pandas as pd
import random
from pathlib import Path
from collections import Counter

class WordleWordLearning:
    def __init__(self, excel_filename="words.xlsx"):
        self.excel_path = Path(__file__).parent / excel_filename
        self.words = []
        self.load_words()
    
    def load_words(self):
        """从Excel文件加载单词和释义（自动转为小写）"""
        try:
            df = pd.read_excel(self.excel_path, header=None)
            self.words = list(zip(df[0].str.lower(), df[1]))  # 加载时直接转为小写
            print(f"已加载 {len(self.words)} 个单词")
        except FileNotFoundError:
            print(f"错误：文件 {self.excel_path} 不存在")
        except Exception as e:
            print(f"读取Excel文件时出错：{e}")
    
    def select_word(self):
        """随机选择一个单词"""
        return random.choice(self.words) if self.words else (None, None)
    
    def check_guess(self, guess, target):
        """检查猜测并返回反馈（✓○×）"""
        feedback = ['×'] * len(target)
        remaining = Counter(target)
        
        # 第一遍扫描处理正确位置
        for i, (g, t) in enumerate(zip(guess, target)):
            if g == t:
                feedback[i] = '✓'
                remaining[g] -= 1
        
        # 第二遍扫描处理存在但位置错误
        for i, (g, t) in enumerate(zip(guess, target)):
            if feedback[i] == '✓':
                continue
            if remaining.get(g, 0) > 0:
                feedback[i] = '○'
                remaining[g] -= 1
        
        return feedback
    
    def _get_valid_input(self, prompt, validator, error_msg):
        """通用输入验证函数"""
        while True:
            user_input = input(prompt).strip().lower()
            if validator(user_input):
                return user_input
            print(error_msg)
    
    def _play_round(self, target_word, target_meaning):
        """单次游戏回合"""
        print(f"\n开始猜测一个 {len(target_word)} 字母的单词")
        
        for attempt in range(6):
            guess = self._get_valid_input(
                f"第 {attempt+1} 次尝试（还剩 {6 - attempt} 次机会）：",
                lambda x: x.isalpha() and len(x) == len(target_word),
                f"请输入有效的 {len(target_word)} 字母单词"
            )
            
            feedback = self.check_guess(guess, target_word)
            print(" ".join(feedback))
            
            if all(mark == '✓' for mark in feedback):
                print(f"恭喜！你猜对了！单词是：{target_word.upper()} - {target_meaning}")
                return
        
        print(f"很遗憾，正确答案是：{target_word.upper()} - {target_meaning}")
    
    def start_game(self):
        """游戏主流程"""
        if not self.words:
            print("无法开始游戏，没有加载到任何单词")
            return
        
        print("欢迎来到Wordle单词背诵游戏！")
        print("规则：你有6次机会猜中一个单词")
        print("✓: 字母位置正确 ○: 字母存在但位置错误 ×: 字母不存在")
        
        while True:
            input("按回车键开始游戏...")
            target_word, target_meaning = self.select_word()
            if not target_word:
                print("错误：未找到可用单词")
                return
            
            self._play_round(target_word, target_meaning)
            
            choice = self._get_valid_input(
                "是否继续游戏？(y/n): ",
                lambda x: x in {'y', 'n'},
                "请输入 y 或 n"
            )
            
            if choice == 'n':
                print("感谢使用，再见！")
                break

if __name__ == "__main__":
    WordleWordLearning().start_game()