# series_allocation.py
from collections import defaultdict
import random
import re
import copy
import math


class SeriesAllocator:
    def __init__(self, custom_rooms=None):
        # Default room capacity configuration
        self.default_room_capacity = {
            "Room1": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                      "Right3": 7},
            "Room2": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                      "Right3": 7},
            "Room3": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6,
                      "Right3": 6},
            "Room4": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                      "Right3": 7},
            "Room5": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                      "Right3": 7},
            "Room6": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                      "Right3": 7},
            "Room7": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                      "Right3": 7},
            "Room8": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6,
                      "Right3": 6},
            "Room9": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                      "Right3": 7},
            "Room10": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room11": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room12": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room13": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6,
                       "Right3": 6},
            "Room14": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room15": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room16": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room17": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room18": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6,
                       "Right3": 6},
            "Room19": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room20": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room21": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room22": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room23": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6,
                       "Right3": 6},
            "Room24": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room25": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room26": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room27": {"Left1": 7, "Left2": 7, "Left3": 7, "Middle1": 6, "Middle2": 6, "Right1": 7, "Right2": 7,
                       "Right3": 7},
            "Room28": {"Left1": 6, "Left2": 6, "Left3": 6, "Middle1": 6, "Middle2": 6, "Right1": 6, "Right2": 6,
                       "Right3": 6},
        }

        # Use custom rooms if provided
        if custom_rooms:
            self.room_capacity = {}
            for room_name, capacity in custom_rooms.items():
                internal_name = room_name.replace(" ", "")
                if isinstance(capacity, dict):
                    self.room_capacity[internal_name] = capacity
                else:
                    self.room_capacity[internal_name] = self._create_block_capacity(capacity)
        else:
            self.room_capacity = self.default_room_capacity

        self.block_order = ["Left1", "Left2", "Left3", "Middle1", "Middle2", "Right1", "Right2", "Right3"]

        # Force exactly 3-4 subjects per room
        self.MIN_SUBJECTS_PER_ROOM = 3
        self.MAX_SUBJECTS_PER_ROOM = 5

        self.MAX_ATTEMPTS = 1000
        self.SMALL_ROOM_THRESHOLD = 10

        # Class grouping for adjacency rules
        self.class_group = {
            "S7CS1": "S7CS", "S7CS2": "S7CS",
            "S7IT": "S7CS",
            "S7EC": "S7ECER", "S7ER": "S7ECER",
            "S7CE": "S7ECER",
            "S5CS1": "S5CS", "S5CS2": "S5CS",
            "S5IT": "S5CS",
            "S5EC": "S5ECE", "S5ER": "S5ECE", "S5CE": "S5ECE",
            "S3CS1": "S3CS", "S3CS2": "S3CS", "S3CS3": "S3CS",
            "S3IT": "S3CS",
            "S3EC": "S3ECER", "S3ER": "S3ECER", "S3CE": "S3ECER",
            "S7CSE": "S7CS",  # Added for new class naming
            "S5CSE": "S5CS",  # Added for new class naming
            "S3CSE": "S3CS"  # Added for new class naming
        }

    def _create_block_capacity(self, total_capacity):
        """Create block-level capacity from total room capacity"""
        # Default distribution: 7,7,7,6,6,7,7,7
        return {
            "Left1": 7, "Left2": 7, "Left3": 7,
            "Middle1": 6, "Middle2": 6,
            "Right1": 7, "Right2": 7, "Right3": 7
        }

    def get_branch(self, cls):
        if len(cls) >= 4:
            return cls[2:4]
        return ""

    def get_semester(self, cls):
        if len(cls) >= 2 and cls[1].isdigit():
            return int(cls[1])
        return 0

    def get_group(self, cls):
        return self.class_group.get(cls, cls)

    def can_place(self, student_class, blocks_in_room, block_index, room_no):
        """Check if a class can be placed in a specific block"""
        current_sem = self.get_semester(student_class)
        current_group = self.get_group(student_class)
        current_branch = self.get_branch(student_class)
        block_name = self.block_order[block_index]

        # Block-level: only one class per block
        if blocks_in_room[block_name]:
            existing_classes = list(blocks_in_room[block_name].keys())
            if existing_classes and student_class not in existing_classes:
                return False

        # Adjacency rules - check left and right neighbors
        def get_classes_in_block(block_idx):
            if 0 <= block_idx < len(self.block_order):
                return set(blocks_in_room[self.block_order[block_idx]].keys())
            return set()

        # Check left neighbor
        if block_index > 0:
            left_classes = get_classes_in_block(block_index - 1)
            for cls in left_classes:
                if self.get_semester(cls) == current_sem:
                    return False
                if self.get_group(cls) == current_group:
                    return False
                if self.get_branch(cls) == "CS" and current_branch == "CS":
                    return False

        # Check right neighbor
        if block_index < len(self.block_order) - 1:
            right_classes = get_classes_in_block(block_index + 1)
            for cls in right_classes:
                if self.get_semester(cls) == current_sem:
                    return False
                if self.get_group(cls) == current_group:
                    return False
                if self.get_branch(cls) == "CS" and current_branch == "CS":
                    return False

        return True

    def get_room_subjects(self, blocks_in_room):
        """Get set of subjects (classes) in a room"""
        subjects = set()
        for block_data in blocks_in_room.values():
            subjects.update(block_data.keys())
        return subjects

    def calculate_room_occupancy(self, blocks_in_room):
        """Calculate total students in a room"""
        total = 0
        for block_data in blocks_in_room.values():
            total += sum(info["count"] for info in block_data.values())
        return total

    def calculate_block_usage(self, blocks_in_room):
        """Calculate which blocks are used"""
        used_blocks = []
        empty_blocks = []
        for block, block_data in blocks_in_room.items():
            if block_data and sum(info["count"] for info in block_data.values()) > 0:
                used_blocks.append(block)
            else:
                empty_blocks.append(block)
        return used_blocks, empty_blocks

    def calculate_balance_score(self, classrooms, leftovers):
        """Calculate score focusing on balance and zero leftovers"""
        total_leftover = sum(leftovers.values())

        # Heavy penalty for leftovers
        if total_leftover > 0:
            return -1000000 * total_leftover

        occupancies = [self.calculate_room_occupancy(blocks) for blocks in classrooms.values()]

        if not occupancies:
            return -float('inf')

        # Room difference (lower is better)
        room_diff = max(occupancies) - min(occupancies)

        # Score components
        score = 0

        # 1. Balance (primary) - penalty for room difference
        score -= room_diff * 1000

        # 2. Bonus for room difference < 10
        if room_diff < 10:
            score += 50000
        elif room_diff < 15:
            score += 25000
        elif room_diff < 20:
            score += 10000

        # 3. Subject count adherence
        for blocks in classrooms.values():
            subjects = len(self.get_room_subjects(blocks))
            if 3 <= subjects <= 4:
                score += 5000
            elif subjects < 3:
                score -= (3 - subjects) * 2000
            else:
                score -= (subjects - 4) * 2000

        # 4. Block usage
        total_blocks = 0
        used_blocks = 0
        for blocks in classrooms.values():
            used, _ = self.calculate_block_usage(blocks)
            used_blocks += len(used)
            total_blocks += len(self.block_order)

        block_usage_pct = (used_blocks / total_blocks) * 100 if total_blocks > 0 else 0
        score += block_usage_pct * 100

        # 5. Small rooms penalty
        small_rooms = sum(1 for occ in occupancies if occ <= self.SMALL_ROOM_THRESHOLD)
        score -= small_rooms * 20000

        return score

    def allocate_session(self, session_data):
        """Allocate all students with balanced rooms"""
        best_solution = None
        best_score = -float('inf')

        # Sort rooms in natural order
        all_rooms = sorted(self.room_capacity.keys(),
                           key=lambda x: int(x.replace("Room", "")))

        total_students = sum(data['count'] for data in session_data.values())
        total_capacity = sum(sum(self.room_capacity[r].values()) for r in all_rooms)

        if total_students > total_capacity:
            print(f"WARNING: Total students ({total_students}) exceed total capacity ({total_capacity})")
            # Still try to allocate as many as possible

        for attempt in range(self.MAX_ATTEMPTS):
            # Create student list
            all_students = []
            for cls, data in session_data.items():
                all_students.extend([cls] * data['count'])

            # Shuffle for randomness
            random.shuffle(all_students)

            # Calculate target per room (as integer)
            target_per_room = total_students // len(all_rooms)  # Integer division
            remaining = total_students - (target_per_room * len(all_rooms))

            # Initialize
            classrooms = {}
            class_assigned = defaultdict(int)
            student_index = 0

            # First pass: Fill rooms with target occupancy
            for i, room_name in enumerate(all_rooms):
                if student_index >= len(all_students):
                    break

                # Distribute remaining students to early rooms
                room_target = target_per_room + (1 if i < remaining else 0)

                blocks, students_placed = self._fill_room_balanced(
                    room_name, all_students, student_index, room_target, session_data
                )

                if blocks:
                    classrooms[room_name] = blocks
                    student_index += students_placed

                    # Update class assigned counts
                    for block_data in blocks.values():
                        for cls, info in block_data.items():
                            class_assigned[cls] += info['count']

            # Check if all allocated
            total_allocated = sum(class_assigned.values())

            if total_allocated == total_students:
                # Calculate balance score
                score = self.calculate_balance_score(classrooms, {})

                # Get room difference
                occupancies = [self.calculate_room_occupancy(b) for b in classrooms.values()]
                room_diff = max(occupancies) - min(occupancies) if occupancies else 0

                if score > best_score:
                    best_score = score
                    best_solution = (classrooms, {}, class_assigned)

                    print(f"Attempt {attempt}: diff={room_diff}, score={score}")

                    # If we find a well-balanced solution, return it
                    if room_diff < 10:
                        print(f"Found balanced solution with diff={room_diff}")
                        return best_solution

            # If not all allocated, try to distribute remaining students
            elif total_allocated < total_students and student_index < len(all_students):
                # Get remaining students
                remaining_students = all_students[student_index:]

                # Try to add them to existing rooms
                classrooms, class_assigned = self._add_remaining_students(
                    classrooms, remaining_students, session_data, class_assigned
                )

                # Recalculate total
                total_allocated = sum(class_assigned.values())

                if total_allocated == total_students:
                    # Rebalance after adding
                    classrooms = self._rebalance_rooms(classrooms, session_data)

                    score = self.calculate_balance_score(classrooms, {})
                    occupancies = [self.calculate_room_occupancy(b) for b in classrooms.values()]
                    room_diff = max(occupancies) - min(occupancies) if occupancies else 0

                    if score > best_score:
                        best_score = score
                        best_solution = (classrooms, {}, class_assigned)

                        if room_diff < 10:
                            print(f"Found balanced solution after adding: diff={room_diff}")
                            return best_solution

        if best_solution:
            return best_solution

        # Final fallback: Force allocation
        return self._force_allocate_all(session_data, all_rooms)

    def _fill_room_balanced(self, room_name, all_students, start_index, target, session_data):
        """Fill a single room aiming for target occupancy"""
        blocks = {block: {} for block in self.block_order}
        room_subjects = set()
        room_occupancy = 0
        room_capacity = sum(self.room_capacity[room_name].values())

        # Don't exceed room capacity
        room_target = min(target, room_capacity)

        # Fill blocks
        for block_idx, block in enumerate(self.block_order):
            if start_index + room_occupancy >= len(all_students):
                break

            if room_occupancy >= room_target:
                break

            block_capacity = self.room_capacity[room_name][block]
            block_used = 0

            # Get students for this block
            current_idx = start_index + room_occupancy
            while (current_idx < len(all_students) and
                   block_used < block_capacity and
                   room_occupancy < room_target):

                cls = all_students[current_idx]

                # Check subject limit
                if cls not in room_subjects and len(room_subjects) >= self.MAX_SUBJECTS_PER_ROOM:
                    current_idx += 1
                    continue

                # Check if we can place
                if self.can_place(cls, blocks, block_idx, int(room_name.replace("Room", ""))):
                    # Count how many of this class are available consecutively
                    cls_count = 0
                    temp_idx = current_idx
                    while temp_idx < len(all_students) and all_students[temp_idx] == cls:
                        cls_count += 1
                        temp_idx += 1

                    # Place as many as possible
                    space_left = min(block_capacity - block_used, room_target - room_occupancy)
                    assign = min(cls_count, space_left)

                    if assign > 0:
                        if cls in blocks[block]:
                            blocks[block][cls]['count'] += assign
                        else:
                            blocks[block][cls] = {
                                'count': assign,
                                'subject': session_data[cls]['subject']
                            }
                            room_subjects.add(cls)

                        block_used += assign
                        room_occupancy += assign
                        current_idx += assign
                    else:
                        current_idx += 1
                else:
                    current_idx += 1

        # If room has students, return blocks and number placed
        if room_occupancy > 0:
            return blocks, room_occupancy
        return None, 0

    def _add_remaining_students(self, classrooms, remaining_students, session_data, class_assigned):
        """Add remaining students to existing rooms"""
        classrooms_copy = copy.deepcopy(classrooms)
        class_assigned_copy = class_assigned.copy()

        # Group remaining students by class
        remaining_by_class = defaultdict(int)
        for cls in remaining_students:
            remaining_by_class[cls] += 1

        # Try to add to rooms with lowest occupancy first
        room_occupancies = [(room, self.calculate_room_occupancy(blocks))
                            for room, blocks in classrooms_copy.items()]
        room_occupancies.sort(key=lambda x: x[1])  # Lowest first

        for cls, count in remaining_by_class.items():
            remaining = count

            for room_name, current_occ in room_occupancies:
                if remaining <= 0:
                    break

                blocks = classrooms_copy[room_name]
                room_subjects = self.get_room_subjects(blocks)

                # Try each block
                for block_idx, block in enumerate(self.block_order):
                    if remaining <= 0:
                        break

                    current = sum(info['count'] for info in blocks[block].values())
                    capacity = self.room_capacity[room_name][block]
                    available = capacity - current

                    if available <= 0:
                        continue

                    # Check if we can place here
                    if cls not in room_subjects and len(room_subjects) >= self.MAX_SUBJECTS_PER_ROOM:
                        continue

                    if self.can_place(cls, blocks, block_idx, int(room_name.replace("Room", ""))):
                        assign = min(remaining, available)

                        if cls in blocks[block]:
                            blocks[block][cls]['count'] += assign
                        else:
                            blocks[block][cls] = {
                                'count': assign,
                                'subject': session_data[cls]['subject']
                            }
                            room_subjects.add(cls)

                        remaining -= assign
                        class_assigned_copy[cls] += assign

        return classrooms_copy, class_assigned_copy

    def _rebalance_rooms(self, classrooms, session_data):
        """Rebalance students between rooms"""
        classrooms_copy = copy.deepcopy(classrooms)

        # Multiple rebalancing passes
        for _ in range(5):
            # Calculate occupancies
            occupancies = [(room, self.calculate_room_occupancy(blocks))
                           for room, blocks in classrooms_copy.items()]

            room_diff = max(o[1] for o in occupancies) - min(o[1] for o in occupancies)

            if room_diff <= 10:
                break

            # Find most and least filled
            most_full = max(occupancies, key=lambda x: x[1])
            least_full = min(occupancies, key=lambda x: x[1])

            if most_full[1] - least_full[1] <= 10:
                break

            # Try to move one student
            moved = self._move_one_student(
                classrooms_copy, most_full[0], least_full[0], session_data
            )

            if not moved:
                break

        return classrooms_copy

    def _move_one_student(self, classrooms, from_room, to_room, session_data):
        """Move one student from one room to another"""
        from_blocks = classrooms[from_room]
        to_blocks = classrooms[to_room]

        # Find a student to move
        for block_idx, block in enumerate(self.block_order):
            if from_blocks[block]:
                cls = list(from_blocks[block].keys())[0]

                # Check if we can place in target room
                for target_idx, target_block in enumerate(self.block_order):
                    current = sum(info['count'] for info in to_blocks[target_block].values())
                    capacity = self.room_capacity[to_room][target_block]

                    if current < capacity:
                        room_subjects = self.get_room_subjects(to_blocks)

                        if cls not in room_subjects and len(room_subjects) >= self.MAX_SUBJECTS_PER_ROOM:
                            continue

                        if self.can_place(cls, to_blocks, target_idx,
                                          int(to_room.replace("Room", ""))):
                            # Move the student
                            from_blocks[block][cls]['count'] -= 1
                            if from_blocks[block][cls]['count'] == 0:
                                del from_blocks[block][cls]

                            if cls in to_blocks[target_block]:
                                to_blocks[target_block][cls]['count'] += 1
                            else:
                                to_blocks[target_block][cls] = {
                                    'count': 1,
                                    'subject': session_data[cls]['subject']
                                }

                            return True

        return False

    def _force_allocate_all(self, session_data, all_rooms):
        """Force allocate all students as a fallback"""
        # Create student list
        all_students = []
        for cls, data in session_data.items():
            all_students.extend([cls] * data['count'])

        # Sort by class to keep together
        all_students.sort()

        classrooms = {}
        class_assigned = defaultdict(int)
        student_index = 0

        for room_name in all_rooms:
            if student_index >= len(all_students):
                break

            blocks = {block: {} for block in self.block_order}
            capacity_map = self.room_capacity[room_name].copy()

            for block in self.block_order:
                if student_index >= len(all_students):
                    break

                block_capacity = capacity_map[block]
                if block_capacity <= 0:
                    continue

                # Get current class
                current_cls = all_students[student_index]

                # Count how many of this class
                count = 0
                for i in range(student_index, len(all_students)):
                    if all_students[i] == current_cls:
                        count += 1
                    else:
                        break

                assign = min(count, block_capacity)

                if assign > 0:
                    blocks[block][current_cls] = {
                        'count': assign,
                        'subject': session_data[current_cls]['subject']
                    }

                    class_assigned[current_cls] += assign
                    student_index += assign
                    capacity_map[block] -= assign

            classrooms[room_name] = blocks

        # Verify all allocated
        total_allocated = sum(class_assigned.values())
        total_needed = sum(data['count'] for data in session_data.values())

        if total_allocated != total_needed:
            print(f"ERROR: Fallback allocation mismatch: {total_allocated}/{total_needed}")

        return classrooms, {}, class_assigned

    def generate_classwise_table(self, class_assigned_count, classrooms):
        """Generate class-wise allocation table with continuous roll numbers"""
        classwise_table = defaultdict(list)

        # First, collect all allocations per class
        class_allocations = defaultdict(list)
        for room, blocks in classrooms.items():
            for block, block_dict in blocks.items():
                for cls, info in block_dict.items():
                    count = info["count"]
                    subject = info["subject"]
                    class_allocations[cls].append({
                        "room": room,
                        "block": block,
                        "count": count,
                        "subject": subject
                    })

        # Sort allocations by room number then block
        for cls in class_allocations:
            class_allocations[cls].sort(key=lambda x: (
                int(x["room"].replace("Room", "")),
                self.block_order.index(x["block"])
            ))

        # Assign continuous roll numbers
        for cls, allocations in class_allocations.items():
            start = 1
            for alloc in allocations:
                count = alloc["count"]
                end = start + count - 1
                classwise_table[cls].append({
                    "Roll Start": start,
                    "Roll End": end,
                    "Room": alloc["room"].replace("Room", "Room "),
                    "Block": alloc["block"],
                    "Subject": alloc["subject"]
                })
                start = end + 1

        return classwise_table

    def generate_series_arrangement(self, series_data):
        """Generate complete seating arrangement"""
        session_data = {}
        for class_item in series_data.get('classes', []):
            class_name = class_item['name']
            strength = class_item['strength']
            subject = series_data.get('subjects', {}).get(class_name, 'Unknown')
            if strength > 0:
                session_data[class_name] = {'count': strength, 'subject': subject}

        print(f"Total students to allocate: {sum(d['count'] for d in session_data.values())}")

        # Run allocation
        classrooms, leftovers, class_assigned = self.allocate_session(session_data)

        # Generate classwise table
        classwise_table = self.generate_classwise_table(class_assigned, classrooms)

        # Convert to results format
        return self.convert_to_results_format(classrooms, session_data, classwise_table, leftovers)

    def convert_to_results_format(self, classrooms, session_data, classwise_table, leftovers):
        """Convert allocation to results format"""
        rooms = {}
        qp_data = []
        subject_totals = defaultdict(int)

        # Track global roll numbers per class
        class_roll_counter = {cls: 1 for cls in session_data.keys()}

        # Sort rooms naturally
        sorted_rooms = sorted(classrooms.keys(),
                              key=lambda x: int(x.replace("Room", "")))

        for room_name in sorted_rooms:
            blocks = classrooms[room_name]

            # Skip empty rooms
            if self.calculate_room_occupancy(blocks) == 0:
                continue

            room_match = re.search(r'Room(\d+)', room_name)
            template_room_name = f"Room {room_match.group(1)}" if room_match else room_name

            room_blocks = {}
            room_subjects = set()
            room_total = 0
            room_capacity_total = 0

            # Process each block
            for block_name, block_data in blocks.items():
                block_capacity = self.room_capacity[room_name][block_name]
                block_students = ['--'] * block_capacity
                block_count = 0

                for cls, info in block_data.items():
                    count = info['count']
                    subject = info['subject']

                    # Assign continuous roll numbers
                    start_roll = class_roll_counter[cls]
                    for i in range(count):
                        roll_number = start_roll + i
                        block_students[i] = f"{cls}:{roll_number}"

                    class_roll_counter[cls] += count

                    block_count += count
                    room_total += count
                    room_capacity_total += block_capacity
                    room_subjects.add(f"{cls}: {subject}")
                    subject_totals[subject] += count

                room_blocks[block_name] = {
                    'students': block_students,
                    'count': block_count,
                    'capacity': block_capacity
                }

            rooms[template_room_name] = {
                'total': room_total,
                'capacity': room_capacity_total,
                'blocks': room_blocks,
                'subjects': list(room_subjects),
                'subjects_count': len(room_subjects)
            }

        # Calculate statistics
        total_students = sum(s['count'] for s in session_data.values())
        allocated_students = sum(r['total'] for r in rooms.values())
        total_capacity = sum(r['capacity'] for r in rooms.values())

        subjects_per_room = [r['subjects_count'] for r in rooms.values()]
        room_occupancies = [r['total'] for r in rooms.values()]

        # Calculate block usage statistics
        total_blocks = 0
        used_blocks = 0
        for room_data in rooms.values():
            for block_data in room_data['blocks'].values():
                total_blocks += 1
                if block_data['count'] > 0:
                    used_blocks += 1

        block_usage = (used_blocks / total_blocks * 100) if total_blocks > 0 else 0

        stats = {
            'total_rooms': len(rooms),
            'total_students': total_students,
            'allocated_students': allocated_students,
            'overflow': total_students - allocated_students,
            'total_capacity': total_capacity,
            'utilization': round((allocated_students / total_capacity * 100), 1) if total_capacity > 0 else 0,
            'avg_occupancy': round(sum(room_occupancies) / len(room_occupancies), 1) if room_occupancies else 0,
            'avg_subjects_per_room': round(sum(subjects_per_room) / len(subjects_per_room),
                                           1) if subjects_per_room else 0,
            'min_subjects': min(subjects_per_room) if subjects_per_room else 0,
            'max_subjects': max(subjects_per_room) if subjects_per_room else 0,
            'room_difference': max(room_occupancies) - min(room_occupancies) if room_occupancies else 0,
            'block_usage': round(block_usage, 1),
            'small_rooms': sum(1 for occ in room_occupancies if occ <= 10)
        }

        qp_summary = {
            'room_wise': qp_data,
            'subject_summary': dict(subject_totals),
            'total_students': sum(subject_totals.values())
        }

        print(f"Allocated {allocated_students} students in {len(rooms)} rooms")
        print(f"Room difference: {stats['room_difference']}")

        return {
            'rooms': rooms,
            'qp_summary': qp_summary,
            'stats': stats,
            'classwise_table': dict(classwise_table),
            'leftovers': leftovers
        }