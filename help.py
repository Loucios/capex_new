class HelpClass:
    @classmethod
    def binary_search(cls, value: float,
                      array: list[object], name: str) -> object:
        right = len(array) - 1
        left = 0
        while right - left != 1:
            center = (left + right) // 2
            if getattr(array[center], name) == value:
                return array[center]
            elif value > getattr(array[center], name):
                left = center
            else:
                right = center
        return (
            array[left] if getattr(array[right], name) - value >
            value - getattr(array[left], name) else array[right]
        )
