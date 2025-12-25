package blackjack

// ParseCard returns the integer value of a card following blackjack ruleset.
func ParseCard(card string) int {
	switch card{
    case "ace":
        return 11
    case  "two":
        return 2
    case "three":
        return 3
    case "four":
        return 4
    case "five":
        return 5
    case "six":
        return 6
    case "seven":
        return 7
    case  "eight":
        return 8
    case  "nine":
        return 9
    case  "ten":
        return 10
    case  "jack":
		return 10
    case "queen":
        return 10
    case "king":
        return 10
    default:
        return 0
    }
}

// FirstTurn returns the decision for the first turn, given two cards of the
// player and one card of the dealer.
func FirstTurn(card1, card2, dealerCard string) string {
	playerVal1 := ParseCard(card1)
	playerVal2 := ParseCard(card2)
	dealerVal := ParseCard(dealerCard)
	playerSum := playerVal1 + playerVal2

	// 策略 1: 如果是一对 Ace，必须分牌 (Split)
	if card1 == "ace" && card2 == "ace" {
		return "P"
	}

	// 策略 2: 如果玩家是 Blackjack (21点)
	if playerSum == 21 {
		// 如果庄家不是 Ace (11) 或 10点牌 (10, J, Q, K)，玩家自动获胜 (Win)
		// 逻辑上就是：如果庄家 < 10，玩家赢；如果庄家 >= 10，玩家停牌 (Stand)
		if dealerVal < 10 {
			return "W"
		}
		return "S"
	}

	// 策略 3: 玩家点数在 [17, 20] 之间，始终停牌 (Stand)
	if playerSum >= 17 && playerSum <= 20 {
		return "S"
	}

	// 策略 4: 玩家点数在 [12, 16] 之间
	if playerSum >= 12 && playerSum <= 16 {
		// 如果庄家是 7 或更高，玩家要牌 (Hit)
		if dealerVal >= 7 {
			return "H"
		}
		// 否则停牌 (Stand)
		return "S"
	}

	// 策略 5: 玩家点数 <= 11，始终要牌 (Hit)
	if playerSum <= 11 {
		return "H"
	}

	// 兜底返回，理论上上面的逻辑覆盖了所有情况
	return "S"
}
