from collections import defaultdict
import random
import re
import copy
import math
from itertools import permutations


class SeriesAllocator:
    def __init__(self, custom_rooms=None):
        self.default_room_capacity = {
            "Room1":  {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room2":  {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room3":  {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6, "Right3": 6},
            "Room4":  {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room5":  {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room6":  {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room7":  {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 2, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room8":  {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6, "Right3": 6},
            "Room9":  {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room10": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room11": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room12": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room13": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6, "Right3": 6},
            "Room14": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room15": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room16": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room17": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room18": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6, "Right3": 6},
            "Room19": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room20": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room21": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room22": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room23": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6, "Right3": 6},
            "Room24": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room25": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room26": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room27": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7, "Right3": 7},
            "Room28": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6, "Right3": 6},
        }

        if custom_rooms:
            self.room_capacity = {}
            for room_name, capacity in custom_rooms.items():
                # Normalise key: "Room 1" → "Room1"
                internal_name = room_name.replace(" ", "")

                if isinstance(capacity, dict):
                    # ── NEW PATH ──────────────────────────────────────────────
                    # Frontend sends the exact per-block dict, e.g.:
                    #   {"Left1":7,"Left2":7,"Left3":7,
                    #    "Middle1":6,"Middle2":6,
                    #    "Right1":7,"Right2":7,"Right3":7}
                    # Validate all 8 expected keys are present; fall back to
                    # the value of the first key in each group if any are missing.
                    expected_keys = [
                        "Left1","Left2","Left3",
                        "Middle1","Middle2",
                        "Right1","Right2","Right3"
                    ]
                    if all(k in capacity for k in expected_keys):
                        # Use exactly what was sent
                        self.room_capacity[internal_name] = {
                            k: int(capacity[k]) for k in expected_keys
                        }
                    else:
                        # Partial dict — infer from whatever keys exist
                        left  = capacity.get("Left1")  or capacity.get("Left2")  or capacity.get("Left3")  or 7
                        mid   = capacity.get("Middle1") or capacity.get("Middle2") or 6
                        right = capacity.get("Right1") or capacity.get("Right2")  or capacity.get("Right3") or 7
                        self.room_capacity[internal_name] = {
                            "Left1": int(left),  "Left2": int(left),  "Left3": int(left),
                            "Middle1": int(mid), "Middle2": int(mid),
                            "Right1": int(right),"Right2": int(right),"Right3": int(right),
                        }
                else:
                    # ── LEGACY PATH ───────────────────────────────────────────
                    # If an integer was passed (old behaviour), distribute it
                    # proportionally across the standard 8-block structure.
                    # Default split: Left=7, Mid=6, Right=7 (same as most rooms).
                    self.room_capacity[internal_name] = self._create_block_capacity(int(capacity))
        else:
            self.room_capacity = self.default_room_capacity

        self.block_order = ["Left1", "Left2", "Left3", "Middle1", "Middle2", "Right1", "Right2", "Right3"]

        self.MIN_SUBJECTS_PER_ROOM = 3
        self.MAX_SUBJECTS_PER_ROOM = 3
        self.MAX_ATTEMPTS = 200
        self.SMALL_ROOM_THRESHOLD = 10

        self.class_group = {
            "S7CS1": "S7CS", "S7CS2": "S7CS", "S7CS3": "S7CS", "S7IT": "S7CS",
            "S7EC": "S7ECER", "S7ER": "S7ECER", "S7CE": "S7ECER",
            "S5CS1": "S5CS", "S5CS2": "S5CS", "S5CS3": "S5CS", "S5IT": "S5CS",
            "S5EC": "S5ECE", "S5ER": "S5ECE", "S5CE": "S5ECE",
            "S3CS1": "S3CS", "S3CS2": "S3CS", "S3CS3": "S3CS", "S3IT": "S3CS",
            "S3EC": "S3ECER", "S3ER": "S3ECER", "S3CE": "S3ECER",
        }

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _create_block_capacity(self, total_capacity):
        """
        Legacy helper: given a single integer total, distribute across 8 blocks
        using the standard 7-7-7-6-6-7-7-7 pattern (total = 52).
        If the requested total differs, we scale proportionally but keep the
        Left/Right uniform and Middle uniform.
        """
        # Standard ratio: Left cols each = 7, Mid cols each = 6, Right cols each = 7
        # 3*left + 2*mid + 3*right = total  with left==right and mid = left-1
        # Solve: 8*left - 2 = total  → left = (total+2)/8
        left = max(1, round((total_capacity + 2) / 8))
        mid  = max(1, left - 1)
        return {
            "Left1": left,  "Left2": left,  "Left3": left,
            "Middle1": mid, "Middle2": mid,
            "Right1": left, "Right2": left, "Right3": left,
        }

    def get_branch(self, cls):
        match = re.match(r'S\d+([A-Z]+)', cls)
        if match:
            branch = match.group(1)
            return re.sub(r'\d+$', '', branch)
        return ""

    def get_semester(self, cls):
        match = re.match(r'S(\d+)', cls)
        if match:
            return int(match.group(1))
        return 0

    def get_group(self, cls):
        return self.class_group.get(cls, cls)

    def _same_sem_same_branch(self, cls_a, cls_b):
        if self.get_semester(cls_a) != self.get_semester(cls_b):
            return False
        branch_a = self.get_branch(cls_a)
        branch_b = self.get_branch(cls_b)
        if branch_a == branch_b:
            return True
        ec_family = {"EC", "ER", "CE"}
        if branch_a in ec_family and branch_b in ec_family:
            return True
        cs_family = {"CS", "IT"}
        if branch_a in cs_family and branch_b in cs_family:
            return True
        return False

    def get_room_subjects(self, blocks_in_room):
        subjects = set()
        for block_data in blocks_in_room.values():
            subjects.update(block_data.keys())
        return subjects

    def room_subject_count(self, blocks_in_room):
        return len(self.get_room_subjects(blocks_in_room))

    def can_add_class_to_room(self, cls, blocks_in_room, strict=True):
        existing = self.get_room_subjects(blocks_in_room)
        if cls in existing:
            return True
        limit = self.MAX_SUBJECTS_PER_ROOM if strict else self.MAX_SUBJECTS_PER_ROOM + 1
        return len(existing) < limit

    def can_place(self, student_class, blocks_in_room, block_index, room_no, strict_subject_limit=True):
        block_name = self.block_order[block_index]
        if blocks_in_room[block_name]:
            existing = list(blocks_in_room[block_name].keys())
            if existing and student_class not in existing:
                return False
        if not self.can_add_class_to_room(student_class, blocks_in_room, strict=strict_subject_limit):
            return False

        def classes_at(idx):
            if 0 <= idx < len(self.block_order):
                return set(blocks_in_room[self.block_order[idx]].keys())
            return set()

        for nbr in classes_at(block_index - 1) | classes_at(block_index + 1):
            if self._same_sem_same_branch(nbr, student_class):
                return False
        return True

    def calculate_room_occupancy(self, blocks_in_room):
        return sum(
            info["count"]
            for block_data in blocks_in_room.values()
            for info in block_data.values()
        )

    def calculate_block_usage(self, blocks_in_room):
        used  = [b for b, d in blocks_in_room.items() if d]
        empty = [b for b, d in blocks_in_room.items() if not d]
        return used, empty

    # ------------------------------------------------------------------
    # Scoring
    # ------------------------------------------------------------------

    def calculate_score(self, classrooms, leftovers):
        if sum(leftovers.values()) > 0:
            return -float('inf')
        score = 0
        occupancies = []
        subject_counts = []
        used_blocks = 0
        total_blocks = 0

        for room_name, blocks in classrooms.items():
            used, _ = self.calculate_block_usage(blocks)
            used_blocks  += len(used)
            total_blocks += len(self.block_order)
            occ = self.calculate_room_occupancy(blocks)
            occupancies.append(occ)
            sc = self.room_subject_count(blocks)
            subject_counts.append(sc)
            if sc < self.MIN_SUBJECTS_PER_ROOM:
                score -= 10000  * (self.MIN_SUBJECTS_PER_ROOM - sc)
            if sc > self.MAX_SUBJECTS_PER_ROOM:
                score -= 100000 * (sc - self.MAX_SUBJECTS_PER_ROOM)

        score += sum(occupancies) * 100
        block_usage_pct = (used_blocks / total_blocks * 100) if total_blocks else 0
        score += block_usage_pct * 100

        if occupancies:
            avg = sum(occupancies) / len(occupancies)
            variance = sum((o - avg) ** 2 for o in occupancies) / len(occupancies)
            score -= variance * 0.5

        return score

    # ------------------------------------------------------------------
    # Linear ordering helpers
    # ------------------------------------------------------------------

    def _valid_linear_order(self, class_list):
        n = len(class_list)
        if n == 0: return []
        if n == 1: return list(class_list)

        by_semester = defaultdict(list)
        for cls in class_list:
            by_semester[self.get_semester(cls)].append(cls)

        semesters = sorted(by_semester.keys())
        result = []
        while any(by_semester.values()):
            for sem in semesters:
                if by_semester[sem]:
                    result.append(by_semester[sem].pop(0))

        for i in range(len(result) - 1):
            if self._same_sem_same_branch(result[i], result[i + 1]):
                limit = min(720, math.factorial(n))
                for j, perm in enumerate(permutations(class_list)):
                    if j >= limit:
                        break
                    if all(not self._same_sem_same_branch(perm[k], perm[k + 1])
                           for k in range(len(perm) - 1)):
                        return list(perm)
                return None
        return result

    # ------------------------------------------------------------------
    # Core allocation
    # ------------------------------------------------------------------

    def _choose_classes_for_room(self, active_sorted, session_data, class_assigned,
                                  target_subjects, strict=True):
        candidates = active_sorted[:min(target_subjects * 3, len(active_sorted))]
        by_semester = defaultdict(list)
        for cls in candidates:
            by_semester[self.get_semester(cls)].append(cls)

        semesters = sorted(by_semester.keys())
        chosen = []
        temp = {s: list(v) for s, v in by_semester.items()}
        while len(chosen) < target_subjects and any(temp.values()):
            for sem in semesters:
                if len(chosen) >= target_subjects:
                    break
                if temp[sem]:
                    chosen.append(temp[sem].pop(0))

        if len(chosen) < target_subjects:
            return None
        order = self._valid_linear_order(chosen)
        if order is None:
            return None
        return order

    def allocate_with_no_leftovers(self, session_data, strict_subject_limit=True):
        all_rooms = sorted(
            self.room_capacity.keys(),
            key=lambda x: int(x.replace("Room", ""))
        )
        total_students = sum(d['count'] for d in session_data.values())

        classrooms = {}
        class_assigned = defaultdict(int)

        for room_idx, room_name in enumerate(all_rooms):
            total_allocated = sum(class_assigned.values())
            if total_allocated >= total_students:
                break

            room_cap = sum(self.room_capacity[room_name].values())
            rooms_left = len(all_rooms) - room_idx
            students_left = total_students - total_allocated
            room_target = min(room_cap, math.ceil(students_left / rooms_left))

            active = [
                cls for cls in session_data
                if session_data[cls]['count'] - class_assigned[cls] > 0
            ]
            if not active:
                break

            active_sorted = sorted(active, key=lambda c: -(session_data[c]['count'] - class_assigned[c]))

            num_available = len(active_sorted)
            target_subjects = min(self.MAX_SUBJECTS_PER_ROOM, max(self.MIN_SUBJECTS_PER_ROOM, num_available))

            chosen_order = None
            for ts in range(target_subjects, self.MIN_SUBJECTS_PER_ROOM - 1, -1):
                chosen_order = self._choose_classes_for_room(
                    active_sorted, session_data, class_assigned, ts, strict=strict_subject_limit
                )
                if chosen_order:
                    break

            if not chosen_order:
                chosen_order = active_sorted[:min(self.MAX_SUBJECTS_PER_ROOM, len(active_sorted))]

            blocks = {b: {} for b in self.block_order}
            room_occ = 0

            remaining_by_class = {
                cls: session_data[cls]['count'] - class_assigned[cls]
                for cls in chosen_order
            }

            for bi, block_name in enumerate(self.block_order):
                if room_occ >= room_target:
                    break

                block_cap = self.room_capacity[room_name][block_name]
                available_cap = block_cap - sum(info['count'] for info in blocks[block_name].values())
                if available_cap <= 0:
                    continue

                best_class = None
                best_count = 0

                for cls in chosen_order:
                    if remaining_by_class.get(cls, 0) <= 0:
                        continue
                    if self.can_place(cls, blocks, bi, int(room_name.replace("Room", "")),
                                      strict_subject_limit=strict_subject_limit):
                        if remaining_by_class[cls] > best_count:
                            best_count = remaining_by_class[cls]
                            best_class = cls

                if best_class:
                    assign = min(available_cap, remaining_by_class[best_class], room_target - room_occ)
                    if assign > 0:
                        if best_class in blocks[block_name]:
                            blocks[block_name][best_class]['count'] += assign
                        else:
                            blocks[block_name][best_class] = {
                                'count': assign,
                                'subject': session_data[best_class]['subject'],
                            }
                        class_assigned[best_class] += assign
                        remaining_by_class[best_class] -= assign
                        room_occ += assign

            if room_occ > 0:
                classrooms[room_name] = blocks

        total_allocated = sum(class_assigned.values())
        if total_allocated < total_students:
            self._place_remaining_students(
                session_data, classrooms, class_assigned, all_rooms,
                strict_subject_limit=strict_subject_limit
            )

        total_allocated = sum(class_assigned.values())
        if total_allocated != total_students:
            print(f"WARNING: Allocated {total_allocated}/{total_students}")

        return classrooms, {}, class_assigned

    def _place_remaining_students(self, session_data, classrooms, class_assigned, all_rooms,
                                   strict_subject_limit=True):
        for cls in session_data:
            remaining = session_data[cls]['count'] - class_assigned[cls]
            if remaining <= 0:
                continue

            for room_name in list(classrooms.keys()):
                if remaining <= 0:
                    break
                blocks = classrooms[room_name]
                if cls not in self.get_room_subjects(blocks):
                    continue
                remaining = self._fill_class_into_room(
                    cls, remaining, blocks, room_name, session_data, class_assigned,
                    strict_subject_limit=strict_subject_limit
                )

            if remaining <= 0:
                continue

            for room_name in all_rooms:
                if remaining <= 0:
                    break
                if room_name not in classrooms:
                    classrooms[room_name] = {b: {} for b in self.block_order}
                blocks = classrooms[room_name]
                if not self.can_add_class_to_room(cls, blocks, strict=strict_subject_limit):
                    continue
                remaining = self._fill_class_into_room(
                    cls, remaining, blocks, room_name, session_data, class_assigned,
                    strict_subject_limit=strict_subject_limit
                )

            if remaining > 0:
                print(f"  FALLBACK (ignore subject limit) for {cls}: {remaining} students")
                for room_name in all_rooms:
                    if remaining <= 0:
                        break
                    if room_name not in classrooms:
                        classrooms[room_name] = {b: {} for b in self.block_order}
                    blocks = classrooms[room_name]
                    remaining = self._fill_class_into_room(
                        cls, remaining, blocks, room_name, session_data, class_assigned,
                        strict_subject_limit=False
                    )

    def _fill_class_into_room(self, cls, remaining, blocks, room_name,
                               session_data, class_assigned, strict_subject_limit=True):
        for bi, block in enumerate(self.block_order):
            if remaining <= 0:
                break
            block_used = sum(info['count'] for info in blocks[block].values())
            block_cap  = self.room_capacity[room_name][block]
            available  = block_cap - block_used
            if available <= 0:
                continue
            if self.can_place(cls, blocks, bi, int(room_name.replace("Room", "")),
                              strict_subject_limit=strict_subject_limit):
                assign = min(remaining, available)
                if assign > 0:
                    if cls in blocks[block]:
                        blocks[block][cls]['count'] += assign
                    else:
                        blocks[block][cls] = {
                            'count': assign,
                            'subject': session_data[cls]['subject'],
                        }
                    class_assigned[cls] += assign
                    remaining -= assign
        return remaining

    # ------------------------------------------------------------------
    # Post-processing
    # ------------------------------------------------------------------

    def post_process_allocation(self, classrooms, session_data):
        classrooms_copy = copy.deepcopy(classrooms)
        for room_name, blocks in classrooms_copy.items():
            for empty_block in [b for b in self.block_order if not blocks[b]]:
                block_idx = self.block_order.index(empty_block)
                for source_block in self.block_order:
                    if source_block == empty_block or not blocks[source_block]:
                        continue
                    for cls, info in list(blocks[source_block].items()):
                        if info['count'] > 1:
                            if self.can_place(cls, blocks, block_idx,
                                              int(room_name.replace("Room", "")),
                                              strict_subject_limit=True):
                                info['count'] -= 1
                                blocks[empty_block][cls] = {
                                    'count': 1,
                                    'subject': session_data[cls]['subject'],
                                }
                                break
                    if blocks[empty_block]:
                        break
        return classrooms_copy

    # ------------------------------------------------------------------
    # Session allocation
    # ------------------------------------------------------------------

    def allocate_session(self, session_data):
        all_classes = list(session_data.keys())
        best_solution = None
        best_score = -float('inf')

        for attempt in range(self.MAX_ATTEMPTS):
            random.shuffle(all_classes)
            shuffled_data = {cls: session_data[cls] for cls in all_classes}

            classrooms, leftovers, class_assigned = self.allocate_with_no_leftovers(
                shuffled_data, strict_subject_limit=True
            )
            if not classrooms:
                continue

            classrooms = self.post_process_allocation(classrooms, shuffled_data)
            score = self.calculate_score(classrooms, {})

            if score > best_score:
                best_score = score
                best_solution = (classrooms, {}, class_assigned)
                occupancies    = [self.calculate_room_occupancy(b) for b in classrooms.values()]
                subject_counts = [self.room_subject_count(b) for b in classrooms.values()]
                violations     = sum(1 for sc in subject_counts if sc > self.MAX_SUBJECTS_PER_ROOM)
                print(
                    f"Attempt {attempt}: rooms={len(classrooms)}, "
                    f"avg_occ={sum(occupancies)/len(occupancies):.1f}, "
                    f"subjects=[{min(subject_counts)}-{max(subject_counts)}], "
                    f"violations={violations}, score={score:.0f}"
                )

        if best_solution:
            final_classrooms = best_solution[0]
            violations = [
                (rn, self.room_subject_count(b))
                for rn, b in final_classrooms.items()
                if self.room_subject_count(b) > self.MAX_SUBJECTS_PER_ROOM
            ]
            if violations:
                print(f"WARNING: {len(violations)} rooms exceed subject limit: {violations}")
            return best_solution

        print("WARNING: Fallback to single attempt with relaxed limits")
        classrooms, leftovers, class_assigned = self.allocate_with_no_leftovers(
            session_data, strict_subject_limit=False
        )
        return classrooms, leftovers, class_assigned

    # ------------------------------------------------------------------
    # Output helpers
    # ------------------------------------------------------------------

    def generate_classwise_table(self, class_assigned_count, classrooms):
        classwise_table  = defaultdict(list)
        class_allocations = defaultdict(list)

        for room, blocks in classrooms.items():
            for block, block_dict in blocks.items():
                for cls, info in block_dict.items():
                    class_allocations[cls].append({
                        "room":    room,
                        "block":   block,
                        "count":   info["count"],
                        "subject": info["subject"],
                    })

        for cls in class_allocations:
            class_allocations[cls].sort(key=lambda x: (
                int(x["room"].replace("Room", "")),
                self.block_order.index(x["block"]),
            ))

        for cls, allocations in class_allocations.items():
            start = 1
            for alloc in allocations:
                count = alloc["count"]
                end   = start + count - 1
                classwise_table[cls].append({
                    "Roll Start": start,
                    "Roll End":   end,
                    "Room":       alloc["room"].replace("Room", "Room "),
                    "Block":      alloc["block"],
                    "Subject":    alloc["subject"],
                })
                start = end + 1

        return classwise_table

    def generate_series_arrangement(self, series_data):
        session_data = {}
        for class_item in series_data.get('classes', []):
            class_name = class_item['name']
            strength   = class_item['strength']
            subject    = series_data.get('subjects', {}).get(class_name, 'Unknown')
            if strength > 0:
                session_data[class_name] = {'count': strength, 'subject': subject}

        classrooms, leftovers, class_assigned = self.allocate_session(session_data)
        classwise_table = self.generate_classwise_table(class_assigned, classrooms)
        return self.convert_to_results_format(
            classrooms, session_data, classwise_table, leftovers
        )

    def convert_to_results_format(self, classrooms, session_data, classwise_table, leftovers):
        rooms = {}
        subject_totals    = defaultdict(int)
        class_roll_counter = {cls: 1 for cls in session_data.keys()}

        sorted_rooms = sorted(
            classrooms.keys(), key=lambda x: int(x.replace("Room", ""))
        )

        subjects_per_room = []

        for room_name in sorted_rooms:
            blocks = classrooms[room_name]
            if self.calculate_room_occupancy(blocks) == 0:
                continue

            room_match = re.search(r'Room(\d+)', room_name)
            template_room_name = f"Room {room_match.group(1)}" if room_match else room_name

            room_blocks  = {}
            room_subjects = set()
            room_total   = 0

            for block_name, block_data in blocks.items():
                block_capacity = self.room_capacity[room_name][block_name]
                block_students = ['--'] * block_capacity
                block_count    = 0
                seat_idx       = 0

                for cls, info in block_data.items():
                    count   = info['count']
                    subject = info['subject']
                    start_roll = class_roll_counter[cls]
                    for i in range(count):
                        if seat_idx < block_capacity:
                            block_students[seat_idx] = f"{cls}_{start_roll + i}"
                            seat_idx += 1
                    class_roll_counter[cls] += count
                    block_count  += count
                    room_total   += count
                    room_subjects.add(f"{cls}: {subject}")
                    subject_totals[subject] += count

                room_blocks[block_name] = {
                    'students': block_students,
                    'count':    block_count,
                    'capacity': block_capacity,
                }

            rooms[template_room_name] = {
                'total':          room_total,
                'capacity':       sum(self.room_capacity[room_name].values()),
                'blocks':         room_blocks,
                'subjects':       list(room_subjects),
                'subjects_count': len(room_subjects),
            }
            subjects_per_room.append(len(room_subjects))

        total_students     = sum(s['count'] for s in session_data.values())
        allocated_students = sum(r['total'] for r in rooms.values())
        total_capacity     = sum(r['capacity'] for r in rooms.values())
        room_occupancies   = [r['total'] for r in rooms.values()]

        total_blocks = used_blocks = 0
        for room_data in rooms.values():
            for block_data in room_data['blocks'].values():
                total_blocks += 1
                if block_data['count'] > 0:
                    used_blocks += 1

        if subjects_per_room:
            avg_subjects = sum(subjects_per_room) / len(subjects_per_room)
            min_subjects = min(subjects_per_room)
            max_subjects = max(subjects_per_room)
        else:
            avg_subjects = min_subjects = max_subjects = 0

        violations = sum(1 for sc in subjects_per_room if sc > self.MAX_SUBJECTS_PER_ROOM)

        stats = {
            'total_rooms':            len(rooms),
            'total_students':         total_students,
            'allocated_students':     allocated_students,
            'overflow':               total_students - allocated_students,
            'total_capacity':         total_capacity,
            'utilization':            round((allocated_students / total_capacity * 100), 1) if total_capacity else 0,
            'avg_occupancy':          round(sum(room_occupancies) / len(room_occupancies), 1) if room_occupancies else 0,
            'avg_subjects_per_room':  round(avg_subjects, 1),
            'min_subjects':           min_subjects,
            'max_subjects':           max_subjects,
            'room_difference':        max(room_occupancies) - min(room_occupancies) if room_occupancies else 0,
            'block_usage':            round((used_blocks / total_blocks * 100), 1) if total_blocks else 0,
            'small_rooms':            sum(1 for occ in room_occupancies if occ <= self.SMALL_ROOM_THRESHOLD),
            'subject_limit_violations': violations,
        }

        if subjects_per_room:
            ideal_subjects  = 4
            subject_score   = 100 - min(100, abs(avg_subjects - ideal_subjects) * 20)
            occupancy_score = (allocated_students / total_capacity * 100) if total_capacity else 0
            stats['balance_score'] = round((subject_score * 0.3 + occupancy_score * 0.7), 1)

        return {
            'rooms':      rooms,
            'qp_summary': {
                'room_wise':       [],
                'subject_summary': dict(subject_totals),
                'total_students':  sum(subject_totals.values()),
            },
            'stats':           stats,
            'classwise_table': dict(classwise_table),
            'leftovers':       {},
        }
