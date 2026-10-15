def find_critical_floor(target_floor: int, total_floors: int = 100) -> tuple[int, int]:
    step = 1
    while (step * (step + 1)) // 2 < total_floors:
        step += 1

    current_floor = 0
    drops = 0

    while current_floor < total_floors:
        drops += 1
        next_floor = min(current_floor + step, total_floors)

        if next_floor >= target_floor:
            break

        current_floor = next_floor
        step -= 1


    for floor in range(current_floor + 1, next_floor + 1):
        if floor == next_floor and target_floor > next_floor:

            return total_floors, drops

        if floor == target_floor:
            return floor, drops

        drops += 1

    return current_floor, drops


target = 72
found, total_drops = find_critical_floor(target_floor=target)
print(f"Target Floor: {target}")
print(f"Found Floor : {found}")
print(f"Total Drops : {total_drops}")
