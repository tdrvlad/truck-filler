'''
TODO:
 - export excel with packages / truck
 - implement mass thresshold
 - (opt) alternate positions
'''



from PIL import Image, ImageDraw, ImageFont
import numpy as np
import excel2json
import glob, os, yaml
import pandas

parameter_file = 'parameters.yaml'
param = yaml.load(open(parameter_file), Loader = yaml.FullLoader)

truck_length = param['truck_length']
truck_width = param['truck_width']
space_percentage = param['space_between_boxes_percentage_from_truck_length']

box_identification_head = param['box_identification_column_name']
quantity_head = param['quantity_column_name']
dimensions_head = param['dimensions_column_name']
dimension_l = param['dimension_l_column_name']
dimension_w = param['dimension_w_column_name']
weight_head = param['package_weight_column_name']
max_weight_per_truck = param['max_weight_per_truck']

def isnan(num):
    return num != num

def remove_previous():

    image_files = glob.glob('*.jpg')
    for image in image_files:
        os.remove(image)

    result_files = glob.glob('Result.xlsx')
    for result_file in result_files:
        os.remove(result_file)


class Box:

    def __init__(self, ident, length, width, weight):

        self.ident = ident
        self.length = length
        self.width = width
        if self.length < self.width:
            self.rotate()
        self.weight = weight


    def rotate(self):

        aux = self.length
        self.length = self.width
        self.width = aux


    def place(self, pos_x, pos_y):

        self.pos_x = pos_x
        self.pos_y = pos_y


    def fits(self, available_length, available_width, available_weight):

        if self.weight < available_weight:
            if self.length < available_length and self.width < available_width:
                print('\nBox ({}x{}) fits in ({:.2f}x{:.2f})'.format(self.length, self.width, available_length, available_width))
                return True

            if self.length < available_width and self.width < available_length:
                print('\nBox ({}x{}) fits rotated in ({:.2f}x{:.2f})'.format(self.length, self.width, available_length, available_width))
                self.rotate()
                return True
                
        return False


class Truck:

    def __init__(self):

        self.boxes = []
        self.available_weight = max_weight_per_truck

        self.length = truck_length
        self.width = truck_width

    def add_box(self, box):

        self.boxes.append(box)
        self.available_weight -= box.weight


class TruckFiller:
    
    def __init__(self, boxes):

        self.space = space_percentage * truck_length
        self.trucks = []

        self.boxes = boxes
        self.boxes = sorted(self.boxes, key = lambda b: b.length * b.width)
        self.boxes.reverse()

        self.place()


    def place(self):
        
        while len(self.boxes) > 0:
            
            self.truck = Truck()
            self.place_by_length(
                pos_x = 0 + self.space, 
                pos_y = 0 + self.space, 
                available_length = self.truck.length - 2 * self.space, 
                available_width = self.truck.width - 2 * self.space,
                )
            self.trucks.append(self.truck)
        

    def place_by_length(self, pos_x, pos_y, available_length, available_width, rev = False):
    
        for box in self.boxes:
            
            if box.fits(available_length, available_width, self.truck.available_weight):
                
                # Remove the placed box from the list
                self.boxes.pop(self.boxes.index(box))

                '''
                   rev (reverse) assures that boxes are not placed biassed to one side of the truck
                '''

                if rev == True:

                    box.place(pos_x, pos_y + available_width - box.width)   
                    self.truck.add_box(box)

                    print('L-Placed box {} ({}x{}) at ({},{}) (rev)'.format(box.ident, box.length, box.width,pos_x, pos_y + available_width - box.width))
                
                    self.place_by_width(
                        pos_x = pos_x, 
                        pos_y = pos_y, 
                        available_length = box.length  + self.space,
                        available_width = available_width - box.width - self.space
                    )

                    self.place_by_length(
                        pos_x = pos_x + box.length + self.space,
                        pos_y = pos_y,
                        available_length = available_length - box.length,
                        available_width = available_width
                    )

                else:

                    box.place(pos_x, pos_y)
                    self.truck.add_box(box)
                    
                    print('L-Placed box {} ({}x{}) at ({},{})'.format(box.ident, box.length, box.width,pos_x, pos_y))
                
                    self.place_by_width(
                        pos_x = pos_x, 
                        pos_y = pos_y + box.width + self.space, 
                        available_length = box.length + self.space,
                        available_width = available_width - box.width - self.space
                    )

                    self.place_by_length(
                        pos_x = pos_x + box.length + self.space,
                        pos_y = pos_y,
                        available_length = available_length - box.length,
                        available_width = available_width,
                        rev = True
                    )
                
                return


    def place_by_width(self, pos_x, pos_y, available_length, available_width):
        
        for box in self.boxes:

            if box.fits(available_length, available_width, self.truck.available_weight):

                # Remove the placed box from the list
                self.boxes.pop(self.boxes.index(box))

                print('W-Placed box {} ({}x{}) at ({},{})'.format(box.ident, box.length, box.width,pos_x, pos_y))
                
                box.place(pos_x, pos_y)
                self.truck.add_box(box)

                self.place_by_width(
                    pos_x = pos_x, 
                    pos_y = pos_y + box.width + self.space, 
                    available_length = available_length,
                    available_width = available_width - box.width
                )
                
                self.place_by_length(
                    pos_x = pos_x + box.length + self.space, 
                    pos_y = pos_y, 
                    available_length = available_length - box.length,
                    available_width = box.width
                )

                return


    def plot(self):

        if (len(self.trucks)) > 0:

            ratio = int(1200 / truck_length)

            for truck in self.trucks:
                border = 20
                font = ImageFont.truetype("arial.ttf", 15)
                image = Image.new("RGB", (truck.length * ratio + 2 * border, truck.width * ratio + 2 * border), (255, 255, 255))
                draw = ImageDraw.Draw(image)   

                # Draw the truck
                draw.rectangle([
                    border ,
                    border, 
                    border + truck.length * ratio,
                    border + truck.width * ratio],
                    outline = (0,0,0),
                    width = 3,
                    fill = (150,150,150)
                )

                font = ImageFont.truetype("arial.ttf", 15)
                text = "Truck {} containing {} packages, total weight: {:.2f} lbs.".format(
                    self.trucks.index(truck) + 1, len(truck.boxes), max_weight_per_truck - truck.available_weight)
                draw.text((border, 1), text, fill = (0,0,0), font=font)
            
                for box in truck.boxes:
                    draw.rectangle([
                        border + box.pos_x * ratio,
                        border + box.pos_y * ratio,
                        border + box.pos_x * ratio + box.length * ratio,
                        border + box.pos_y * ratio + box.width * ratio],
                        outline = (0,0,0),
                        width = 2,
                        fill = tuple(np.random.randint(256, size=3))
                    )
                    font = ImageFont.truetype("arial.ttf", 12)
                    text = "{}".format(box.ident)
                    l, w = draw.textsize(text, font=font)
                    draw.text(
                        (
                            border + ratio * (box.pos_x ) + (box.length * ratio - l) / 2, 
                            border + ratio * (box.pos_y ) + (box.width * ratio - w) / 2
                        ),
                        text, fill = (0,0,0), font=font)

                img_name = 'Truck ' + str(self.trucks.index(truck) + 1) + '.jpg'
                image.save(img_name)


class DataLoader:
    
    def __init__(self):
        
        data_files = glob.glob('*.xlsx')
        if len(data_files) > 0:       
            data_file = max(data_files, key=os.path.getctime)
            self.data = pandas.read_excel(data_file, header = [0, 1])
            self.get_relevant_headers()
        else:
            self.data = None
            print('Found no excel file')
    

    def get_relevant_headers(self):

        self.headers = list(self.data)

        try:
            self.package_ident_head = [head for head in self.headers if box_identification_head in head]
            self.quantity_head = [head for head in self.headers if quantity_head in head]
            self.length_head = [head for head in self.headers if dimensions_head in head and dimension_l in head]
            self.width_head = [head for head in self.headers if dimensions_head in head and dimension_w in head]
            self.weight_head = [head for head in self.headers if weight_head in head]
            
            print('Succesfully read Excel file.')
            
        except:
            print('Error reading Excel Columns.')
        

    def get_boxes(self):

        self.boxes = []
        
        idents = self.data.loc[:, self.package_ident_head].to_numpy().reshape(-1)
        quantities = self.data.loc[:, self.quantity_head].to_numpy().reshape(-1)
        lenghts = self.data.loc[:, self.length_head].to_numpy().reshape(-1)
        widths = self.data.loc[:, self.width_head].to_numpy().reshape(-1)
        weights = self.data.loc[:, self.weight_head].to_numpy().reshape(-1)
       
        if len(idents) == len(quantities) == len(lenghts) == len(widths) == len(weights):

            for i in range(len(self.data)):

                try:   
                    ident = idents[i]
                    quantity = quantities[i]
                    length = np.round(lenghts[i] + 0.01, decimals = 2)
                    width = np.round(widths[i] + 0.01, decimals = 2)
                    weight = np.round(weights[i] + 0.01, decimals = 2)

                    if not isnan(ident) and not isnan(quantity) and not isnan(length) and not isnan(width) and not isnan(weight):
                        for _ in range(int(quantity)):
                            box = Box(ident, length, width, weight)
                            self.boxes.append(box)
                except:
                    print('Error in parsing package data.')

        else:
            print("Error in parsing package data.")

        return self.boxes
        

def write_excel(trucks):

    truck_col = []
    package_ident_col = []
    package_length_col = []
    package_width_col = []
    package_weight_col = []
    
    for truck in trucks:
        for box in truck.boxes:
            truck_col.append('Truck {}'.format(trucks.index(truck) + 1))
            package_ident_col.append(box.ident)
            package_length_col.append(box.length)
            package_width_col.append(box.width)
            package_weight_col.append(box.weight)
    
    df = pandas.DataFrame({
        'Truck' : truck_col,
        'Package ID' : package_ident_col,
        'Package Length' : package_length_col,
        'Package Width' : package_width_col,
        'Package Weight' : package_weight_col
    })

    df.to_excel('Result.xlsx')  



if __name__ == '__main__':

    
    remove_previous()

    data_loader = DataLoader()
    boxes = data_loader.get_boxes()

    for box in boxes:
        print('Package {} - size {}x{}'.format(box.ident, box.length, box.width), flush = True)
    
    truck_filler = TruckFiller(boxes)
    truck_filler.plot()

    write_excel(truck_filler.trucks)
    


