# 三元结构理论 案例分析引擎 — v2.2 版本验证工具
# Copyright (C) 2026 李政恒 (Li Zhengheng)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published
# by the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

from analysis_engine import CaseAnalysisEngine

if __name__ == "__main__":
    engine = CaseAnalysisEngine("theory_config.json")
    
    # Material ID and Link
    input_list = ["素材1 - 湖南石门救灾村干部金耳环舆情事件分析报告。"]
    
    # Actual analysis content following v2.2 high-tension guidelines
    analysis_results = {
        "title": "救灾现场的“金耳环”审判：道德代偿与实干叙事的剧烈对冲",
        "core_tension": "一个在泥泞中奋战数日的村支书，因为耳边的一抹金色，在屏幕后被审判为‘精致的腐败者’。——问题是：苦难必须以‘自虐’来证明其真实吗？",
        "abstract": "本报告针对湖南石门县救灾村干部向某某因佩戴金耳环引发的舆情风波进行深度拆解。基于生存+繁衍+逆反的三维度交叉分析，揭示了网络受众通过‘道德洁癖’完成精神代偿的深层机制。分析认为，这种基于符号的道德审判不仅扭曲了治理评价标准，更对基层治理主体的存续动力造成了结构性伤害。",
        "background": "时间/背景：2026年5月，湖南石门强降雨救灾一线。主体：村支书向某某及网络审判者群体。事件：因采访中佩戴金耳环引发网暴，焦点从救灾实绩偏移至个人饰品。",
        "core_problem": "核心矛盾在于‘表演性苦难’的社会期待与‘正常化物质生活’的基层现实之间的断裂。这种断裂被社交媒体放大后，形成了一场基于符号误读的道德狂欢。",
        "po_survival": "生存维度：受众在现实压力下存在精神失衡，倾向于通过审判他人来获得‘正义代偿’。向某某在物质维度已支出极高成本，却因一个社会维度符号导致其生存价值被全盘消解。\n这就是‘符号’的威力。\n它比事实更沉重。",
        "po_reproduction": "繁衍维度：金耳环作为高辨识度符号，在短视频环境下实现了裂变式复制。传播中议题发生严重变异：从‘救灾进展’变成了‘财产来源’的无底洞猜疑。\n真相被淹没了。\n取而代之的是不断自我强化的偏见。",
        "po_reactivity": "逆反维度：这是一种群体自发的破坏性逆反。金耳环触发了网民的‘火药桶效应’，将对现实的不满投射到实干者身上，直接导致基层治理者的逆反失能风险。\n锁死了。全锁死了。\n实干者正变得越来越沉默。",
        "perspective_switch": "如果站在一个普通家长的视角：我更希望看到的是一个能把我从洪水中救出来的支书，还是一个因为没戴耳环而被网民夸赞‘淳朴’但动作慢半拍的符号？答案不言而喻。",
        "po_cross_analysis": "三维度交叉：生存根基驱动逆反，逆反通过变异路径放大。这种‘生存不满+变异扩散+逆反升级’的正向循环，正在摧毁社会的信任根基。",
        "solutions": "1. 叙事重塑：用‘硬核绩效数据’覆盖‘软性道德符号’。 2. 制度保护：建立基层实干者的‘舆论避风港’，拒绝无底洞式的自证清白。",
        "conclusion": "本质机制：在情绪敏感期，个体的物质符号若挑战了受众的精神代偿边界，极易引发破坏性逆反。这种‘符号政治’是对实质正义的极大讽刺。\n金句：当一个社会的英雄必须以‘赤贫’为勋章时，这个社会离真实的繁荣就越来越远。",
        "open_question": "留给中国基层治理的问题是：如果‘金耳环’成了审判实干者的唯一标准，未来还有谁愿意在泥泞中挺身而出？"
    }

    # Mock extraction
    engine.extract_elements = lambda content: {
        "subject": "向某某（救灾村干）",
        "event": "金耳环舆情网暴",
        "background": "2026年湖南石门救灾一线"
    }

    engine.run_batch(input_list, target_dir="v2.2_reports", analysis_results_list=[analysis_results])
