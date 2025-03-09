import pandas as pd
import random
from pathlib import Path
from collections import Counter
import time

class WordleWordLearning:
    def __init__(self, excel_filename="words.xlsx"):
        self.excel_path = Path(__file__).parent / excel_filename
        self.words = []
        self.selected_length = None  # 新增：存储用户选择的单词长度
        self.load_words()
    
    def load_words(self):
        """从Excel文件加载单词和释义"""
        try:
            df = pd.read_excel(self.excel_path, header=None)
            # 加载时直接转为小写并过滤非字母字符
            self.words = [
                (str(w).lower().strip(), m) 
                for w, m in zip(df[0], df[1]) 
                if isinstance(w, str) and w.isalpha()
            ]
            print(f"已加载 {len(self.words)} 个有效单词")
        except FileNotFoundError:
            print(f"错误：文件 {self.excel_path} 不存在")
        except Exception as e:
            print(f"读取Excel文件时出错：{e}")
    
    def get_available_lengths(self):
        """获取可用的单词长度列表"""
        return sorted({len(w) for w, _ in self.words})
    
    def prompt_length_selection(self):
        """用户选择单词长度"""
        available = self.get_available_lengths()
        print("\n可选的单词长度:", available)
        
        return self._get_valid_input(
            prompt="请选择单词长度（0表示不限长度）: ",
            validator=lambda x: x.isdigit() and (int(x) in available or x == "0"),
            error_msg=f"请输入有效的长度选项（0或{available}中的数字）"
        )
    
    def select_word(self):
        """根据选择的长度随机选词"""
        candidates = self.words
        if self.selected_length and self.selected_length != "0":
            candidates = [(w, m) for w, m in self.words if len(w) == int(self.selected_length)]
        
        if not candidates:
            print(f"没有长度为 {self.selected_length} 的可用单词")
            return None, None
        
        return random.choice(candidates)
    
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
    
    def _animate_feedback(self, feedback):
        """动态显示反馈结果"""
        for i, mark in enumerate(feedback):
            # 最后一个符号不添加空格
            end_char = '\n' if i == len(feedback)-1 else ' '
            # 立即刷新输出缓冲区
            print(mark, end=end_char, flush=True)
            if i != len(feedback)-1:
                time.sleep(0.2)  # 符号间间隔0.2秒

    def _play_round(self, target_word, target_meaning):
        """单次游戏回合"""
        print(f"\n目标单词长度：{len(target_word)} 字母")
        
        for attempt in range(6):
            guess = self._get_valid_input(
                prompt=f"第 {attempt+1} 次尝试（还剩 {6 - attempt} 次机会）: ",
                validator=lambda x: x.isalpha() and len(x) == len(target_word),
                error_msg=f"请输入有效的 {len(target_word)} 字母单词"
            )
            
            feedback = self.check_guess(guess, target_word)
            self._animate_feedback(feedback)  # 使用动态显示方法
            
            if all(mark == '✓' for mark in feedback):
                print(f"恭喜！你猜对了！单词是：{target_word.upper()} - {target_meaning}")
                return True
        
        print(f"很遗憾，正确答案是：{target_word.upper()} - {target_meaning}")
        return False
    
    def start_game(self):
        """游戏主流程"""
        if not self.words:
            print("无法开始游戏，没有加载到任何单词")
            return
        
        print("欢迎来到Wordle单词背诵游戏！")
        print("规则：你有6次机会猜中一个单词")
        print("✓: 字母位置正确 ○: 字母存在但位置错误 ×: 字母不存在")
        
        # 初始长度选择
        self.selected_length = self.prompt_length_selection()
        
        while True:
            target_word, target_meaning = self.select_word()
            if not target_word:
                # 处理无可用单词的情况
                if self.selected_length != "0":
                    print("当前长度没有可用单词，请重新选择")
                    self.selected_length = self.prompt_length_selection()
                    continue
                else:
                    print("错误：未找到可用单词")
                    return
            
            self._play_round(target_word, target_meaning)
            
            choice = self._get_valid_input(
                prompt="是否继续游戏？（a-继续当前长度/s-重新选择长度/d-退出）: ",
                validator=lambda x: x in {'a', 's', 'd'},
                error_msg="请输入 a/s/d"
            )
            
            if choice == 'd':
                print("感谢使用，再见！")
                break
            elif choice == 's':
                self.selected_length = self.prompt_length_selection()

if __name__ == "__main__":
    WordleWordLearning().start_game()