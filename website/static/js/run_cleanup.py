from backend import db_ops as ops

print("Before cleanup:")
for row in ops.debug_bad_components():
    print(row)

ops.delete_unlinked_blank_components()

print("\nAfter cleanup:")
for row in ops.debug_bad_components():
    print(row)