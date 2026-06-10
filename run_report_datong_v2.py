# -*- coding: utf-8 -*-
#
# 三元结构理论 案例分析引擎 — 大同订婚案 v2 报告生成器
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

"""生成\"山西大同订婚强奸案\"三元结构分析报告 v2.0（核心命题驱动版）"""
import sys, os
os.chdir(r"D:\360MoveData\Users\马格斯佩斯科夫\Desktop\人类主义\三元结构理论\三元结构理论 案例分析skill")
sys.path.insert(0, os.getcwd())

from analysis_engine import CaseAnalysisEngine

engine = CaseAnalysisEngine(
    r"D:\360MoveData\Users\马格斯佩斯科夫\Desktop\人类主义\三元结构理论\三元结构理论 案例分析skill\theory_config.json",
    export_dir=r"D:\360MoveData\Users\马格斯佩斯科夫\Desktop\人类主义\三元结构理论\三元结构理论 案例分析skill\报告"
)

title = '山西大同"订婚强奸案"——礼法交替期的三重代价'

report_text = r'''# 山西大同"订婚强奸案"——礼法交替期的三重代价

## 分析框架前置说明

> 「一个男人付了10万元彩礼，仍然被判强奸；一个女人报了警，仍然被全网质疑。这场判决让双方都觉得自己输了。」

**本报告核心命题：订婚案的判决不是终点，而是礼法体系双重崩塌的起点——彩礼无法购买同意，法律也无法重建信任。**

本报告选取三个概念，从三个角度分别论证这一命题：

**概念一：利益动线（利益维度）**

> 利益动线指：事件中，各类利益（经济资源、社会资本、人身自由）沿特定路径在各方之间流动的轨迹。

- 支撑核心命题的角度：彩礼—婚姻—性权利之间的「利益预期链条」在报警瞬间发生了不可逆断裂，而断裂后的资源被司法系统重新分配——男女双方都没有拿到自己期望的「完整胜利」，这是「双方都觉得输了」的物质基础。

**概念二：合法性松动（价值维度）**

> 合法性松动指：当旧有规范（传统婚俗）失去社会认同，而新规范（现代法治的性同意标准）尚未扎根底层生活世界时，两套规范同时失效的过渡真空状态。

- 支撑核心命题的角度：彩礼之所以「买不到同意」，是因为传统规范对彩礼功能的定义已经被现代法律彻底推翻——但这种推翻在底层社会的感知中并不被接受，导致「法律赢了官司，但输了人心」，即法律也「无法重建信任」。

**概念三：代偿性撕裂（生存维度）**

> 代偿性撕裂指：在安全感受损的情境下，社会群体以极端认同某一方立场来代偿自身的安全感缺失，这种代偿不会弥合冲突，反而会自激放大对立。

- 支撑核心命题的角度：当「彩礼无法购买同意」和「法律无法重建信任」同时成立时，男女双方都无法从这个判决中找到安全感——于是各自启动代偿模式，撕裂成为系统性输出，而非个案的偶然。

**本报告试图回答的问题：**
为什么一张三年的判决书，让所有人都觉得自己在这个社会里更不安全了？

## 概念解读：订婚在中国的"双重面孔"

在进入分析前，有必要说清「订婚」这一制度在中国社会的双重身份——

**面孔一：传统礼俗中的"准婚姻"**

山西农村至今沿用「看地方」习俗：男方家庭带彩礼登门，女方宗亲实地查验，交钱后女儿「留家」视为默许。这套逻辑源自《礼记》「纳征既成，中道不弃」——在传统语境里，彩礼不是「购物」，而是两个家族之间的信用凭证，彩礼交付后两人之间的关系在社会意义上已经成立。

**面孔二：现代法律中的"零效力仪式"**

中国法律不承认订婚产生任何夫妻权利——不共有财产，不产生扶养义务，更不构成性同意的推定。订婚在法律面前是一个民间仪式，效力为零。

> 本案的核心冲突，就是这两副面孔在同一张床上的正面碰撞。

## 案例事实摘要

### 一、当事人与时间线

| 时间 | 事件 |
|------|------|
| 2023年1月 | 席某某与吴某某经婚介所相识，确立恋爱关系 |
| 2023年5月1日 | 举办订婚仪式，男方支付10万元彩礼及戒指，签订协议「婚后一年房产加女方名」 |
| 2023年5月2日 | 在婚房发生性关系，女方随后逃至13楼呼救报警，指控强奸 |
| 2023年5月5日 | 席某某被刑事拘留 |
| 2023年12月25日 | 一审：强奸罪成立，判处有期徒刑三年 |
| 2025年4月16日 | 二审：维持原判，民事部分彩礼无需返还 |

### 二、核心争议

> 医院检查显示女方「处女膜完整」，送检物未检出精斑。

法院认定：强奸罪的核心标准是「是否违背意志」，不以身体损伤为唯一依据。依女方陈述、逃离监控、及时报警行为及情绪崩溃等间接证据，形成完整证据链。

男方家属质疑：「未等DNA结果即批捕」，已提交对12名办案人员的追责申请。

自被拘至二审宣判，席某某已羁押712天。

## 三元结构分析正文

### 一、彩礼链条的断裂：谁拿走了什么

```DIAGRAM
{"viz":"network","title":"彩礼链条断裂：利益动线前后对比","nodes":[{"id":"groom_f","label":"男方家庭","type":"material","detail":"输出方：10万彩礼+戒指+房产加名期权，预期回报为婚姻关系确立"},{"id":"bride_f","label":"女方家庭","type":"material","detail":"获得方：10万现金+戒指+房产期权，传统逻辑下为婚约信用凭证"},{"id":"marriage","label":"婚姻预期","type":"identity_culture","detail":"传统逻辑的预期回报：婚姻关系确立，含传统语境中性关系的默许"},{"id":"alarm","label":"报警触发","type":"security","detail":"断裂点：女方报警，法律逻辑瞬间接管，传统信用链条崩溃"},{"id":"civil","label":"民事判决","type":"institutional_future","detail":"彩礼无需返还，锁定在女方——传统信用的物质载体被法律固化"},{"id":"criminal","label":"刑事判决","type":"political","detail":"3年有期徒刑——人身自由归国家，刑罚执行"},{"id":"opinion","label":"舆论碾压","type":"public","detail":"社会信任蒸发：双方名誉均受损，任何一方都无法取回公信力"},{"id":"result_m","label":"男方结局","type":"material","detail":"财失+人失：彩礼不退+3年刑期，传统预期的全部回报归零"},{"id":"result_f","label":"女方结局","type":"security","detail":"财保+名誉损：保住彩礼，但被全网叙事攻击为「诬告者」「设套者」"}],"edges":[{"source":"groom_f","target":"bride_f","label":"10万彩礼+戒指+房产期权","type":"economic","detail":"传统逻辑：彩礼=信用凭证，交换婚姻关系确立"},{"source":"bride_f","target":"marriage","label":"传统默许","type":"cultural","detail":"彩礼交付后，女方家庭默许婚姻关系（含传统语境的性关系默认）"},{"source":"marriage","target":"groom_f","label":"预期回报","type":"economic","detail":"男方家庭的全部预期：婚姻确立=付出获得回报"},{"source":"alarm","target":"civil","label":"法律接管","type":"legal","detail":"报警后民事逻辑启动：彩礼归属由法院裁定"},{"source":"alarm","target":"criminal","label":"法律接管","type":"legal","detail":"报警后刑事逻辑启动：强奸罪认定与量刑"},{"source":"alarm","target":"opinion","label":"舆论引爆","type":"cultural","detail":"案件公开后舆论场撕裂，性别对立叙事生成"},{"source":"civil","target":"result_f","label":"彩礼锁定","type":"economic","detail":"民事判决：彩礼无需返还，女方保住经济利益"},{"source":"criminal","target":"result_m","label":"自由剥夺","type":"legal","detail":"刑事判决：3年有期徒刑，男方失去人身自由"},{"source":"opinion","target":"result_m","label":"名誉攻击","type":"cultural","detail":"男方被标签化，但更多舆论同情"},{"source":"opinion","target":"result_f","label":"名誉攻击","type":"cultural","detail":"女方被全网质疑为「诬告者」「为彩礼设套」"}]}
```

本案最容易被误读的地方，是把「彩礼判归女方」解读为「女方赢了」。

实际上，利益动线的断裂制造了一个罕见的三败局面：男方失去了经济筹码和人身自由；女方拿到了10万彩礼，却同时承受了「诬告者」「为彩礼设套」的全网叙事攻击；司法系统完成了裁决，却在民间积累了大量「法律偏袒女性」的不信任账单。三方各取一部分「胜利」，同时承担一部分「损失」，没有人能宣称「我赢了」。

这个结构本身就是利益动线断裂的最直接证明——断裂前，彩礼是男方的经济杠杆，预期回报是一段完整的婚姻关系；断裂后，彩礼变成了司法判决的证据物，回报变成了刑期和舆论烂账。**预期的「利益交换」变成了实际的「利益蒸发」。**

→ **彩礼链条的断裂不是「女方赢了」，而是「三方都输了」——这就是核心命题「彩礼无法购买同意」的物质现实。**

### 二、两套规范的同时失效：为什么法律赢了官司，却输了人心

利益动线告诉我们彩礼最终去了哪里——但它没有回答更根本的问题：为什么男方家庭在付出彩礼后会「理所当然」地认为发生性关系是被允许的？法律判决为什么不能平息这种认知？这正是**合法性松动**要揭开的结构性根源。

```DIAGRAM
{"viz":"network","title":"两套规范的碰撞：传统礼俗vs现代法治","nodes":[{"id":"tradition","label":"传统礼俗逻辑","type":"identity_culture","detail":"纳征既成，留家即默许——彩礼是两个家族间的信用凭证，交付后关系在社会意义上已成立"},{"id":"modern_law","label":"现代法律逻辑","type":"institutional_future","detail":"订婚零法律效力——性同意须持续明确，任何关系不免责"},{"id":"naizheng","label":"纳征既成·留家即默许","type":"cultural","detail":"传统规范核心：彩礼=信用凭证，交付后性关系被视为默许"},{"id":"consent","label":"性同意须持续明确","type":"legal","detail":"现代法治核心：任何关系不构成性同意推定，必须持续明确"},{"id":"collide","label":"规范碰撞带","type":"security","detail":"同一张床上的正面碰撞——传统逻辑说「已经默许」，法律逻辑说「从未同意」"},{"id":"trad_decay","label":"传统礼俗溃败","type":"identity_culture","detail":"被自身物质化掏空神圣性——订婚已成「彩礼谈判桌+房产加名磋商」，礼俗道义外壳被利益交换腐蚀"},{"id":"law_decay","label":"现代法治溃败","type":"political","detail":"赢了判决，失了民心——底层没打过官司不信法律有用，判决结果让两边都觉得法律偏袒"},{"id":"vacuum","label":"礼法双不信真空","type":"public","detail":"两套规范同时失效：人们既不再信礼，也还没学会信法——规范真空地带"}],"edges":[{"source":"tradition","target":"naizheng","label":"核心信条","type":"cultural","detail":"彩礼交付=信用凭证=性关系默许"},{"source":"modern_law","target":"consent","label":"核心原则","type":"legal","detail":"订婚不产生任何法律效力，性同意必须持续明确"},{"source":"naizheng","target":"collide","label":"传统逻辑","type":"cultural","detail":"「已经默许」——彩礼交付后的关系认定"},{"source":"consent","target":"collide","label":"法律逻辑","type":"legal","detail":"「从未同意」——法律对性自主权的绝对保护"},{"source":"collide","target":"trad_decay","label":"传统侧溃败","type":"cultural","detail":"礼俗被自身商品化从内部腐蚀——不是被打倒，而是自己掏空"},{"source":"collide","target":"law_decay","label":"法律侧溃败","type":"legal","detail":"判决正确≠判决被接受——底层世界对法律的合法性尚未建立"},{"source":"trad_decay","target":"vacuum","label":"信任崩塌","type":"cultural","detail":"传统规范失效，但文化惯性仍在运行"},{"source":"law_decay","target":"vacuum","label":"信任崩塌","type":"legal","detail":"现代法治未被底层接受，判决无法终止传统逻辑"}]}
```

合法性松动最典型的症状，是其中没有哪一方「赢了」——两套规范都在本案中暴露出各自的脆弱性。

传统礼俗的脆弱性来自内部：当订婚仪式已经演变为「房产加名协议签订现场」，传统礼俗的神圣外衣早在男方拿出协议那一刻就被自己撕破了。这不是女性意识觉醒「打倒了」传统礼俗，而是传统礼俗被自身的商品化逻辑从内部腐蚀掉了。

现代法律的脆弱性来自底层接受度：判决确实在法律逻辑上完整——但在山西农村，「绝大多数人一生打不了一次官司」，法律对他们的意义是「出了事才用的东西」，而不是「日常行为的框架」。判决的正确不等于判决的被接受。

真正危险的结局不是「新法律战胜了旧礼俗」，而是两者**同时**变得不可信。法院的判决无法终止传统礼俗的逻辑，传统礼俗的失效也无法让人们转向现代法治。底层生活世界出现了规范真空——人们既不再信「礼」，也还没学会信「法」。

→ **这正是核心命题的第二层：法律的判决「无法重建信任」——因为法律本身在底层世界的合法性，也随着礼俗的崩塌一同进入了失效区间。**

### 三、安全感的螺旋分裂：为什么一个案子刺穿了所有人

合法性松动解释了「为什么双方都不服」——但它还没有解释为什么这种「不服」会跨越当事人两家、蔓延成全国性的性别战争。这个放大机制，是**代偿性撕裂**。

```DIAGRAM
{"viz":"network","title":"代偿性撕裂：安全感危机的自激放大","nodes":[{"id":"trigger","label":"判决公布","type":"public","detail":"触发事件：二审判决维持原判，三年刑期确定，彩礼无需返还"},{"id":"male_perceive","label":"男性群体感知","type":"security","detail":"彩礼+订婚都不管用，人财两空的风险具象化——传统承诺的安全保障一文不值"},{"id":"female_perceive","label":"女性群体感知","type":"security","detail":"终于有判决证明了性同意不可交易——法律确认了身体自主权"},{"id":"male_comp","label":"男性代偿行为","type":"material","detail":"拒绝高额彩礼、收缩婚恋支出、「婚姻风险太大」——经济收缩策略"},{"id":"female_comp","label":"女性代偿行为","type":"political","detail":"强化「不就是不」、要求更多法律保护——法律防御策略"},{"id":"male_signal","label":"女性读出的信号","type":"security","detail":"「男人更算计了，婚前更危险」——男性的收缩被视为攻击"},{"id":"female_signal","label":"男性读出的信号","type":"security","detail":"「女人随时可以诬告，更危险」——女性的防御被视为威胁"},{"id":"loop","label":"自激循环","type":"public","detail":"每一方的防御行为都成为对方的新威胁信号——螺旋升级，无终局"}],"edges":[{"source":"trigger","target":"male_perceive","label":"男性侧冲击","type":"cultural","detail":"「在男性网民心里狠狠刺了一刀」——卢克文"},{"source":"trigger","target":"female_perceive","label":"女性侧确认","type":"legal","detail":"判决证明「不就是不」不是空洞口号"},{"source":"male_perceive","target":"male_comp","label":"代偿启动","type":"economic","detail":"安全感缺失→经济收缩策略"},{"source":"female_perceive","target":"female_comp","label":"代偿启动","type":"legal","detail":"安全感确认→法律防御策略"},{"source":"male_comp","target":"female_signal","label":"信号误读","type":"cultural","detail":"男性收缩→女性读出「男人更危险」"},{"source":"female_comp","target":"male_signal","label":"信号误读","type":"cultural","detail":"女性坚持→男性读出「随时可以诬告」"},{"source":"male_signal","target":"loop","label":"新威胁","type":"security","detail":"男性感知到新的不安全→更深层收缩"},{"source":"female_signal","target":"loop","label":"新威胁","type":"security","detail":"女性感知到新的不安全→更强力防御"},{"source":"loop","target":"male_perceive","label":"螺旋升级","type":"cultural","detail":"新一轮代偿→更深的收缩","bidirectional":false},{"source":"loop","target":"female_perceive","label":"螺旋升级","type":"cultural","detail":"新一轮代偿→更强的防御","bidirectional":false}]}
```

卢克文说这个案子「在男性网民心里狠狠刺了一刀」。这刀刺的不是案情本身，而是一个延迟觉醒的恐惧——原来传统逻辑承诺的「安全保障」（彩礼换婚姻、付出换回报）在法律面前一文不值。

男性群体的代偿是「收缩」：拒绝高额彩礼、拒绝婚前大额经济承诺。女性群体的代偿是「确认」：这个判决证明了「不就是不」不再是空洞口号。

但两种代偿之间有一个致命的结构性错位：男性的收缩行为，在女性眼里是「男人更算计了，婚前更危险」；女性的坚持行为，在男性眼里是「法律随时可以用来对付男人」。每一方的防御行为都被对方读作新的攻击信号。

这不是信息不对称，不是沟通问题，不是谁更坏——这是一个**安全感的结构性冲突**：在传统规范失效、新规范未扎根的间隙地带，男性和女性维持安全感的方式本身就是互相拆台的。

→ **代偿性撕裂完成了核心命题的闭环：彩礼买不到同意，法律重建不了信任，而这两件事同时成立，使得所有人的安全感都无处安放——这才是「全国性别战争」的真正根源。**

## 结论

### 汇流：三条证据，指向同一个现实

利益动线告诉我们：彩礼链条断裂后，没有任何一方拿到「完整胜利」——男方失去了一切，女方保住了彩礼却失去了名誉，司法系统完成了判决却积累了不信任。

合法性松动告诉我们：这场断裂发生在两套规范**同时**失效的真空地带——传统礼俗被商品化腐蚀，现代法治未被底层接受，任何一套规范都无法单独承担「重建秩序」的任务。

代偿性撕裂告诉我们：在这种规范真空里，男女双方各自启动的安全感防御机制，不是解决问题的出口，而是问题的自我复制装置。

**三者共同指向的，正是本报告的核心命题：订婚案的判决不是终点，而是礼法体系双重崩塌的起点——彩礼无法购买同意，法律也无法重建信任。**

### 核心判断

这个案件之所以「刺穿了所有人」，不是因为它特别极端，而是因为它恰好击中了中国社会当下最脆弱的那条缝隙：婚恋规范的「新旧交替期」。在这个时间窗口里，旧规范（彩礼=信用凭证=性权利默认）已经失去法律效力，但其文化惯性仍然在数千万人的婚恋行为中运行；新规范（性同意必须持续明确）已经写入法律，但其社会根基还远未稳固。本案的当事人，就是在这条缝隙里行动的普通人——他们的悲剧，本质上是两套时间表错位的代价。

判决可以结案，但无法弥合这条缝隙。真正填满它的，只能是一代人的观念重置——而那需要的时间，远比一张判决书更长。

> 判决书裁定了一个人是否有罪，但裁定不了一个时代的规范应该是什么。当「纳征既成」不再灵验，而「不就是不」还没扎进每一块泥土——每一个站上法庭的人，都只是两套规范交替时被碾过的代价。而那条缝隙本身，还在那里。

## 附录

**数据来源**：
- [知乎《山西大同"订婚强奸案"事件详细始末（简洁版）》](https://zhuanlan.zhihu.com/p/1896177587228754747) (2025.04.17)
- [网易《山西订婚强奸案二审维持原判！三年刑期背后的法律逻辑》](https://www.163.com/dy/article/JT8V56DP0556CF0W.html) (2025.04.16)
- [网易《山西大同"订婚强奸案"，二审判了！一纸判决书，撕裂…》](https://www.163.com/dy/article/JT9Q9IJK0521U536.html) (2025.04.16)
- [搜狐《卢克文：山西大同案影响很大，男性网民心里已被"深深刻下一刀"》](https://www.sohu.com/a/887180051_121948378) (2025.04.21)
- [腾讯新闻《从大同"订婚强奸案"看中国社会的秩序撕裂》](https://news.qq.com/rain/a/20250421A03IF200) (2025.04.21)
- [搜狐《不就是不——从大同订婚强奸案读懂强奸罪的法律边界》](https://www.sohu.com/a/1028506276_122700490) (2024)
- 大同市中级人民法院二审判决书（2025.04.16）
- 阳高县人民法院一审判决书（2023.12.25）

**分析框架**：
- 三元结构理论 v5.0（利益动线 / 合法性松动 / 代偿性撕裂）
- 核心命题驱动结构（核心命题→三角论证→汇流结论）
'''

# 直接调用 export_from_text，传入完整的 Markdown 报告文本
result = engine.export_from_text(title, report_text)

if isinstance(result, dict):
    print(f"\n=== 报告生成完成 ===")
    print(f"文件夹: {result['folder']}")
    print(f"Word: {result['word']}")
    print(f"PDF: {result['pdf']}")
else:
    print(f"PDF: {result}")
